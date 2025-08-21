package app.sme.util;

import org.apache.poi.xwpf.usermodel.*;

import java.util.Map;

public class WordUtil {
    public static void replacePlaceHolder(XWPFParagraph paragraph, Map<String, String> map) {
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.text();
            boolean changed = false;
            for (Map.Entry<String, String> e : map.entrySet()) {
                if (text.contains(e.getKey())) {
                    text = text.replace(e.getKey(), e.getValue());
                    changed = true;
                }
            }
            if (changed) {
                run.setText(text, 0);
            }
        }
    }
    public static void replaceTextInDocument(XWPFDocument doc, Map<String, String> map) {

        // Replace trong các đoạn văn bản thường
        for (XWPFParagraph p : doc.getParagraphs()) {
            replacePlaceHolder(p, map);
        }

        // Replace trong bảng (nếu tài liệu có table)
        for (XWPFTable table : doc.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        replacePlaceHolder(p, map);
                    }
                }
            }
        }

        // Replace trong header
        for (XWPFHeader header : doc.getHeaderList()) {
            for (XWPFParagraph p : header.getParagraphs()) {
                replacePlaceHolder(p, map);
            }
            for (XWPFTable table : header.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            replacePlaceHolder(p, map);
                        }
                    }
                }
            }
        }

        // Replace trong footer (nếu muốn thêm luôn)
        for (XWPFFooter footer : doc.getFooterList()) {
            for (XWPFParagraph p : footer.getParagraphs()) {
                replacePlaceHolder(p, map);
            }
            for (XWPFTable table : footer.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            replacePlaceHolder(p, map);
                        }
                    }
                }
            }
        }
    }

}
