package ru.nsu.fit.g14201.marchenko;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;

import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 *
 */
public class DateFilterApp {
    public static void main(String[] args){
        String sexOne = null, sexTwo = null;
        String[] sexOptions = {"Мужчина", "Женщина"};

        String windowName = "Text filter for vk.com";


        JOptionPane.showMessageDialog(null,
                "Памятка пользователю:\n" +
                        "- Только для обновлённой версии vk!\n" +
                        "- Перед использованием удалить из текста все смайлики, отображающиеся в Word как квадраты. " +
                        "Некоторые смайлы преображаются корректно, но знать наверняка, как программа их преобразует, нельзя.\n" +
                        "- Программа не рассчитана на корректную обработку репостов, осторожнее с ними.\n" +
                        "- Шрифт первой строки скорее всего не совпадает со шрифтом остального текста, но это чертовски мелкий недостаток.\n" +
                        "- Преобразованный документ находится в той же директории, что и данная программа, и имеет название \"It's better, dude.docx\" ",
                windowName, JOptionPane.PLAIN_MESSAGE);

        String docName = JOptionPane.showInputDialog(null, "Введите имя файла (без расширения)." + "\n" +
                "Внимание! Программа работает только с файлами, созданными в Microsoft Word 2007 и новее.", windowName, JOptionPane.PLAIN_MESSAGE);
        if (docName == null || docName.isEmpty()) {
            JOptionPane.showMessageDialog(null, "Имя файла не было введено. Программа прекращает работу.", windowName, JOptionPane.ERROR_MESSAGE);
            System.exit(0);
        }

        String nameOne = JOptionPane.showInputDialog(null, "Введите имя первого собеседника.", windowName, JOptionPane.QUESTION_MESSAGE);
        if (nameOne == null || nameOne.isEmpty()) {
            JOptionPane.showMessageDialog(null, "Имя первого собеседника не было введено. Программа прекращает работу.", windowName, JOptionPane.ERROR_MESSAGE);
            System.exit(0);
        }

        int firstGender = JOptionPane.showOptionDialog(null,
                "Какой пол у персонажа, которого зовут " + nameOne + "?",
                windowName,
                JOptionPane.DEFAULT_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                sexOptions,
                "Мужчина");
        switch (firstGender) {
            case JOptionPane.CLOSED_OPTION:
                JOptionPane.showMessageDialog(null, "Пол первого собеседника не выбран. Программа прекращает работу.", windowName, JOptionPane.ERROR_MESSAGE);
                System.exit(0);
                break;
            case 1:
                sexOne = "Ж";
                break;
            default:
                sexOne = "М";
                break;
        }

        String nameTwo = JOptionPane.showInputDialog(null, "Введите имя второго собеседника.", windowName, JOptionPane.QUESTION_MESSAGE);
        if (nameTwo == null || nameTwo.isEmpty()) {
            JOptionPane.showMessageDialog(null, "Имя первого собеседника не было введено. Программа прекращает работу.", windowName, JOptionPane.ERROR_MESSAGE);
            System.exit(0);
        }

        int secondGender = JOptionPane.showOptionDialog(null,
                "Какой пол у персонажа, которого зовут " + nameTwo + "?",
                windowName,
                JOptionPane.DEFAULT_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                sexOptions,
                "Мужчина");
        switch (secondGender) {
            case JOptionPane.CLOSED_OPTION:
                JOptionPane.showMessageDialog(null, "Пол второго собеседника не выбран. Программа прекращает работу.", windowName, JOptionPane.ERROR_MESSAGE);
                System.exit(0);
                break;
            case 1:
                sexTwo = "Ж";
                break;
            default:
                sexTwo = "М";
                break;
        }

        Filter filter = new Filter(nameOne, nameTwo, sexOne, sexTwo);

        try (FileOutputStream out = new FileOutputStream(new File("It's better, dude.docx"))) {
            OPCPackage opcPackage = OPCPackage.open(docName + ".docx");

            filter.process(out, opcPackage, windowName);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Didn't manage to create XWPFDocument.", windowName, JOptionPane.ERROR_MESSAGE);
            e.printStackTrace();
        } catch (InvalidFormatException | IllegalStateException e) {
            JOptionPane.showMessageDialog(null, "File hasn't been opened.", windowName, JOptionPane.ERROR_MESSAGE);
            e.printStackTrace();
        }
    }
}