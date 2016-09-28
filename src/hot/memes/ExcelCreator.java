package hot.memes;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColorScaleFormatting;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import java.awt.AlphaComposite;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import javax.imageio.ImageIO;
import org.apache.commons.io.FilenameUtils;

public class ExcelCreator {

    double scale = 100;

    public void run(File file, int scale) {
        this.scale = scale / 100.0;
        createExcel(file);
    }

    public static String toAlphabetic(int i) {

        int quot = i / 26;
        int rem = i % 26;
        char letter = (char) ((int) 'A' + rem);
        if (quot == 0) {
            return "" + letter;
        } else {
            return toAlphabetic(quot - 1) + letter;
        }
    }

    public void createExcel(File file) {

        InputStream in;
        int width = 0;
        int height = 0;
        boolean resize = false;

        try {
            BufferedImage original = ImageIO.read(file);
            double IMG_WIDTH = original.getWidth();
            double IMG_HEIGHT = original.getHeight();

            if (scale != 0) {
                int newWidth = (int) (IMG_WIDTH * scale);
                BufferedImage resizedImage = new BufferedImage(newWidth, (int) ((newWidth / IMG_WIDTH) * IMG_HEIGHT), original.getType());
                Graphics2D g = resizedImage.createGraphics();
                g.drawImage(original, 0, 0, newWidth, (int) ((newWidth / IMG_WIDTH) * IMG_HEIGHT), null);
                g.dispose();
                g.setComposite(AlphaComposite.Src);
                g.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
                g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
                g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

                ImageIO.write(resizedImage, "bmp", new File(FilenameUtils.removeExtension(file.getAbsolutePath()) + "_resized.bmp"));

                resize = true;
            }
        } catch (IOException ex) {
            Logger.getLogger(ExcelCreator.class.getName()).log(Level.SEVERE, null, ex);
        }

        byte[] pixelData;

        try {
            byte[] b = new byte[4];
            in = resize ? new FileInputStream(FilenameUtils.removeExtension(file.getAbsolutePath()) + "_resized.bmp") : new FileInputStream(file);
            in.skip(18);
            in.read(b);
            for (int i = 0; i < 4; i++) {
                width += (b[i] & 0xFF) * Math.pow(256, i);
            }
            in.read(b);
            for (int i = 0; i < 4; i++) {
                height += (b[i] & 0xFF) * Math.pow(256, i);
            }
            in.skip(28);

            int padding = 4 - ((width * 3) % 4) == 4 ? 0 : 4 - ((width * 3) % 4);
            int pwidth = (width * 3) + padding;

            pixelData = new byte[height * pwidth];
            in.read(pixelData);
            XSSFWorkbook wb = new XSSFWorkbook();
            FileOutputStream fileOut = new FileOutputStream(FilenameUtils.removeExtension(file.getAbsolutePath()) + "_resized.bmp" + ".xlsx");
            XSSFSheet sheet = wb.createSheet("New Sheet");

            int index = 0;
            for (int r = height - 1; r >= 0; r--) {
                Row row = sheet.createRow(r);
                for (int c = 0; c < 3 * width; c += 3) {
                    row.createCell(c + 2).setCellValue(pixelData[index++] & 0xFF);
                    row.createCell(c + 1).setCellValue(pixelData[index++] & 0xFF);
                    row.createCell(c).setCellValue(pixelData[index++] & 0xFF);
                }
                index += padding;
                row.setHeight((short) 1000);
            }
            for (int i = 0; i < width * 3; i++) {
                sheet.setColumnWidth(i, 700);
            }

            XSSFSheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

            XSSFConditionalFormattingRule redRule
                    = sheetCF.createConditionalFormattingColorScaleRule();
            XSSFColorScaleFormatting cs1 = redRule.getColorScaleFormatting();
            cs1.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
            cs1.getThresholds()[0].setValue(0.0);
            cs1.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
            cs1.getThresholds()[1].setValue(255.1);
            ((ExtendedColor) cs1.getColors()[0]).setARGBHex("FF000000");
            ((ExtendedColor) cs1.getColors()[1]).setARGBHex("FFFF0000");

            ConditionalFormattingRule greenRule
                    = sheetCF.createConditionalFormattingColorScaleRule();
            ColorScaleFormatting cs2 = greenRule.getColorScaleFormatting();
            cs2.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
            cs2.getThresholds()[0].setValue(0.0);
            cs2.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
            cs2.getThresholds()[1].setValue(255.1);
            ((ExtendedColor) cs2.getColors()[0]).setARGBHex("FF000000");
            ((ExtendedColor) cs2.getColors()[1]).setARGBHex("FF00FF00");

            XSSFConditionalFormattingRule blueRule
                    = sheetCF.createConditionalFormattingColorScaleRule();
            XSSFColorScaleFormatting cs3 = blueRule.getColorScaleFormatting();
            cs3.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
            cs3.getThresholds()[0].setValue(0.0);
            cs3.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.NUMBER);
            cs3.getThresholds()[1].setValue(255.1);
            ((ExtendedColor) cs3.getColors()[0]).setARGBHex("FF000000");
            ((ExtendedColor) cs3.getColors()[1]).setARGBHex("FF0000FF");

            CellRangeAddress[] range;
            for (int i = 0; i < width * 3; i += 3) {
                range = new CellRangeAddress[]{CellRangeAddress.valueOf(toAlphabetic(i) + "1:" + toAlphabetic(i) + height)};
                sheetCF.addConditionalFormatting(range, redRule);
                range = new CellRangeAddress[]{CellRangeAddress.valueOf(toAlphabetic(i + 1) + "1:" + toAlphabetic(i + 1) + height)};
                sheetCF.addConditionalFormatting(range, greenRule);
                range = new CellRangeAddress[]{CellRangeAddress.valueOf(toAlphabetic(i + 2) + "1:" + toAlphabetic(i + 2) + height)};
                sheetCF.addConditionalFormatting(range, blueRule);
            }

            wb.write(fileOut);
            fileOut.close();
            in.close();

        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelCreator.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelCreator.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
