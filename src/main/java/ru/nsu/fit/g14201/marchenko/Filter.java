package ru.nsu.fit.g14201.marchenko;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.swing.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 *
 */
public class Filter {
    private XWPFDocument newDoc;

    private String nameOne;
    private String nameTwo;
    private String sexOne;
    private String sexTwo;

    private int fontSize;
    private String fontFamily;
    private boolean bold;
    private boolean italic;
    private final int DEFAULT_FONT_SIZE = 10;

    private XWPFParagraph baseParagraph = null;
    private boolean firstTime;

    public Filter(String nOne, String nTwo, String sOne, String sTwo) {
        nameOne = nOne;
        nameTwo = nTwo;
        sexOne = sOne;
        sexTwo = sTwo;

        fontSize = 11;
        fontFamily = "Calibri";
        bold = false;
        italic = false;

        firstTime = true;
    }

    public void process(FileOutputStream out, OPCPackage opcPackage, String windowName) {

        try {
            XWPFDocument oldDoc = new XWPFDocument(opcPackage);
            newDoc = new XWPFDocument();

            Iterator<XWPFParagraph> paragraphIterator = oldDoc.getParagraphs().iterator();
            while (paragraphIterator.hasNext()) {
                XWPFParagraph curPar = paragraphIterator.next();
                String parText = curPar.getParagraphText();
                if (!parText.isEmpty()) {
                    if (isNameAndDate(parText)) { //Future "М@" and "Ж@"
                        baseParagraph = newDoc.createParagraph();
                        String gender;
                        gender = parText.contains(nameOne)? sexOne + "@ " : sexTwo + "@ ";
                        insert(gender, baseParagraph);
                        copyAllRunsToAnotherParagraph(paragraphIterator.next());

                    } else //Usual text
                        if (!whetherMarkOutMessage(parText) && !isDate(parText)) {
                            try {
                                XWPFRun tempRun = baseParagraph.createRun();
                                tempRun.setText(" / ");
                                copyAllRunsToAnotherParagraph(curPar);
                            } catch (NullPointerException exc) {
                                JOptionPane.showMessageDialog(null, "Christina was soooo fucking afraid of this exception. Shit.", windowName, JOptionPane.ERROR_MESSAGE);
                                System.exit(0);
                            }
                    }
                }
            }

            newDoc.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public boolean isNameAndDate(String tested) {
        Pattern pattern = Pattern.compile("^(" + nameOne + "|" + nameTwo + ")\\s\\d?\\d\\:\\d{2}\\s*");
        Matcher matcher = pattern.matcher(tested);

        return matcher.matches();
    }
    public boolean whetherMarkOutMessage(String tested){
        Pattern pattern = Pattern.compile("^Выделить сообщение\\s*$");
        Matcher matcher = pattern.matcher(tested);

        Pattern datePat = Pattern.compile("^\\d?\\d (января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s*");
        Matcher dateMatcher = datePat.matcher(tested);

        return (matcher.matches() || dateMatcher.matches());
    }


    public boolean isDate(String tested) {
        Pattern pattern = Pattern.compile("\\s*\\d?\\d\\.\\d{2}\\.\\d{2}\\s*");
        Matcher matcher = pattern.matcher(tested);

        return matcher.matches();
    }

    // Copy all runs from one paragraph to another, keeping the style unchanged
    private void copyAllRunsToAnotherParagraph(XWPFParagraph oldPar) {
        if (firstTime) {
            List<XWPFRun> runs = oldPar.getRuns();
            if (!runs.isEmpty()) {
                XWPFRun firstRun = runs.get(0);
                fontSize = firstRun.getFontSize();
                fontFamily = firstRun.getFontFamily();
                bold = firstRun.isBold();
                italic = firstRun.isItalic();

                firstTime = false;
            }
        }

        for (XWPFRun run : oldPar.getRuns()) {
            String textInRun = run.getText(0);
            if (textInRun == null || textInRun.isEmpty()) {
                continue;
            }

            XWPFRun newRun = baseParagraph.createRun();

            // Copy text
            newRun.setText(textInRun);

            // Apply the same style
            newRun.setFontSize((fontSize == -1) ? DEFAULT_FONT_SIZE : fontSize);
            newRun.setFontFamily(fontFamily);
            newRun.setBold(bold);
            newRun.setItalic(italic);
        }
    }
    private void insert(String toInsert, XWPFParagraph newParagraph) {
        XWPFRun newRun = newParagraph.createRun();
        newRun.setText(toInsert);

        newRun.setFontSize((fontSize == -1) ? DEFAULT_FONT_SIZE : fontSize);
        newRun.setFontFamily(fontFamily);
    }
}
