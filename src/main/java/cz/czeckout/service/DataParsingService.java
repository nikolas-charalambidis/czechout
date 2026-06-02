package cz.czeckout.service;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.checkerframework.checker.nullness.qual.Nullable;

import cz.czeckout.entity.Account;
import cz.czeckout.entity.Address;
import cz.czeckout.entity.Invoice;
import cz.czeckout.entity.Item;
import cz.czeckout.entity.Metadata;
import cz.czeckout.entity.Method;
import cz.czeckout.entity.Party;
import cz.czeckout.entity.PdfData;
import cz.czeckout.parser.EntityRowMapper;
import lombok.NonNull;


public class DataParsingService {

    private enum Section {
        NONE, PARTIES, ADDRESSES, ACCOUNTS, METHODS, VARIABLES
    }

    @NonNull
    public Map<String, PdfData> parseWorkbook(@NonNull final Path path) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(Files.newInputStream(path))) {
            final var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            final var entityRowMapper = new EntityRowMapper(formulaEvaluator);

            final var metadata = parseMetadata(workbook, entityRowMapper);
            final var sheetData = new HashMap<String, PdfData>();

            // Process all sheets except "ROOT"
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                final var sheet = workbook.getSheetAt(i);
                final var sheetName = sheet.getSheetName();
                
                if (!"ROOT".equals(sheetName)) {
                    final var invoices = parseInvoices(workbook, entityRowMapper, metadata, sheetName);
                    if (!invoices.isEmpty()) {
                        sheetData.put(sheetName, new PdfData(invoices, metadata));
                    }
                }
            }

            return sheetData;
        }
    }

    @NonNull
    private Metadata parseMetadata(@NonNull final Workbook workbook,
                                   @NonNull final EntityRowMapper entityRowMapper) {

        final var sheet = workbook.getSheet("ROOT");
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet ROOT not found");
        }

        final var parties = new HashMap<String, Party>();
        final var addresses = new HashMap<String, Address>();
        final var accounts = new HashMap<String, Account>();
        final var methods = new HashMap<String, Method>();
        final var variables = new LinkedHashMap<String, String>();

        var section = Section.NONE;
        final var header = new HashMap<String, Integer>();

        for (var row : sheet) {
            final var marker = entityRowMapper.getCellStringValue(row.getCell(1)); // column B

            if (marker == null || marker.isBlank()) {
                continue;
            }

            // Section detection
            section = switch (marker.trim()) {
                case "PARTIES" -> Section.PARTIES;
                case "ADDRESSES" -> Section.ADDRESSES;
                case "ACCOUNTS" -> Section.ACCOUNTS;
                case "METHODS" -> Section.METHODS;
                case "VARIABLES" -> Section.VARIABLES;
                default -> section;
            };

            if (marker.matches("PARTIES|ADDRESSES|ACCOUNTS|METHODS|VARIABLES")) {
                header.clear();
                continue;
            }

            // Header row
            if (header.isEmpty()) {
                for (var cell : row) {
                    if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING) {
                        header.put(cell.getStringCellValue(), cell.getColumnIndex());
                    }
                }
                continue;
            }

            if (section == Section.VARIABLES) {
                final var key = entityRowMapper.getCellStringValue(row.getCell(header.get("KEY")));
                final var value = entityRowMapper.getCellStringValue(row.getCell(header.get("VALUE")));
                if (key != null) {
                    variables.put(key, value);
                }
                continue;
            }

            // Data rows
            final var id = entityRowMapper.getCellStringValue(row.getCell(header.get("ID")));
            switch (section) {
                case ADDRESSES -> addresses.put(id, entityRowMapper.mapToAddress(row, header));
                case ACCOUNTS -> accounts.put(id, entityRowMapper.mapToAccount(row, header));
                case METHODS -> methods.put(id, entityRowMapper.mapToPaymentMethod(row, header));
                case PARTIES -> parties.put(id, entityRowMapper.mapToParty(row, header));
            }
        }

        final var metadata = new Metadata(parties, addresses, accounts, methods, variables);
        resolveAddressReferences(metadata);
        return metadata;
    }

    private void resolveAddressReferences(@NonNull final Metadata metadata) {
        metadata.getParties().values().forEach(party -> {
            party.setAddress(metadata.getAddresses().get(party.getAddressReference()));
        });
    }

    @NonNull
    private List<Invoice> parseInvoices(@NonNull final Workbook workbook,
                                        @NonNull final EntityRowMapper entityRowMapper,
                                        @NonNull final Metadata metadata,
                                        @Nullable final String sheetName) {

        final var invoices = new ArrayList<Invoice>();
        final var sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            System.out.println("Sheet '" + sheetName + "' not found, skipping...");
            return invoices;
        }
        final var iterator = sheet.rowIterator();

        Invoice currentInvoice = null;
        Map<String, Integer> invoiceHeaderMap = null;
        boolean invoicesAnchorFound = false;
        Row lookahead = null;

        while (true) {
            Row row;
            if (lookahead != null) {
                row = lookahead;
                lookahead = null;
            } else {
                if (!iterator.hasNext()) {
                    break;
                }
                row = iterator.next();
            }

            if (entityRowMapper.isRowEmpty(row)) {
                continue;
            }

            final var firstNonEmptyCell = entityRowMapper.getFirstNonEmptyCell(row);
            final var firstNonEmpty = entityRowMapper.getCellStringValue(firstNonEmptyCell);

            if (!invoicesAnchorFound) {
                if ("INVOICES".equalsIgnoreCase(firstNonEmpty)) {
                    System.out.println("Found invoices sheet at " + firstNonEmptyCell.getAddress().formatAsString());
                    invoicesAnchorFound = true;
                }
                continue;
            }

            if ("INVOICE".equalsIgnoreCase(firstNonEmpty)) {
                System.out.println("Found invoice at " + firstNonEmptyCell.getAddress().formatAsString());
                // This is invoice header row
                invoiceHeaderMap = entityRowMapper.createHeaderMap(row);
                // Next row = first invoice data row
                final var dataRow = entityRowMapper.findNextNonEmptyRow(iterator);
                if (dataRow != null) {
                    currentInvoice = entityRowMapper.mapToInvoice(dataRow, invoiceHeaderMap, metadata);
                    invoices.add(currentInvoice);
                    // Also parse first item from same row
                    final var item = entityRowMapper.mapToItem(dataRow, invoiceHeaderMap);
                    if (item != null) {
                        currentInvoice.getItems().add(item);
                    }
                }
                continue;
            }

            if (currentInvoice != null) {
                // Check if this row is a new invoice data row
                final var possibleInvoiceNo = entityRowMapper.getCellStringValue(row.getCell(invoiceHeaderMap.get("INVOICE")));
                if (possibleInvoiceNo != null && !possibleInvoiceNo.isBlank()) {
                    lookahead = row; // next iteration will treat as new invoice
                    currentInvoice = null;
                    continue;
                }

                // Otherwise, treat as an item row
                final var item = entityRowMapper.mapToItem(row, invoiceHeaderMap);
                if (item != null) {
                    currentInvoice.getItems().add(item);
                }
            }
        }

        // Post-processing.
        invoices.forEach(invoice -> {
            // Align units if null/empty.
            final var defaultItemUnitLength = invoice.getItems().stream()
                .map(Item::getUnit)
                .filter(Objects::nonNull)
                .mapToInt(String::length)
                .max()
                .orElse(0);
            final var defaultItemUnit = "\u00A0".repeat(defaultItemUnitLength);
            invoice.getItems().forEach(item -> {
                if (StringUtils.isBlank(item.getUnit())) {
                    item.setUnit(defaultItemUnit);
                }
            });
        });

        return invoices;
    }
}
