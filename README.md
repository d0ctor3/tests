# tests

// File: src/main/java/com/example/DocxToPdfAsImage.java

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.List;
import javax.imageio.ImageIO;

import org.apache.poi.xwpf.usermodel.*;

import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.PDPageContentStream;

public class DocxToPdfAsImage {

    /**
     * Converts a DOCX file into a PDF where each page is a rendered image of the DOCX content.
     */
    public static void convertDocxToPdfViaImage(File docxFile, File pdfFile) throws Exception {
        FileInputStream fis = new FileInputStream(docxFile);
        XWPFDocument document = new XWPFDocument(fis);

        PDDocument pdfDocument = new PDDocument();

        // Create 1 image = 1 page
        BufferedImage image = renderDocxToImage(document);
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", baos);
        baos.flush();

        PDPage page = new PDPage(PDRectangle.LETTER);
        pdfDocument.addPage(page);

        PDImageXObject pdImage = PDImageXObject.createFromByteArray(pdfDocument, baos.toByteArray(), "docx-image");
        PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, page);
        contentStream.drawImage(pdImage, 0, 0, page.getMediaBox().getWidth(), page.getMediaBox().getHeight());
        contentStream.close();

        pdfDocument.save(pdfFile);
        pdfDocument.close();
        fis.close();
    }

    /**
     * Renders DOCX content into a BufferedImage (manual layout).
     */
    private static BufferedImage renderDocxToImage(XWPFDocument document) {
        int width = 612; // 8.5in * 72
        int height = 792; // 11in * 72

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = image.createGraphics();

        g.setColor(Color.WHITE);
        g.fillRect(0, 0, width, height);
        g.setColor(Color.BLACK);
        g.setFont(new Font("Serif", Font.PLAIN, 14));

        FontMetrics fm = g.getFontMetrics();
        int x = 50;
        int y = 50;

        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            String text = paragraph.getText();
            if (text != null && !text.trim().isEmpty()) {
                g.drawString(text, x, y);
                y += fm.getHeight();
            }

            // Page break simulation (simplified, optional)
            if (y >= height - 50) {
                break; // this sample only does one page
            }
        }

        g.dispose();
        return image;
    }

    // Example usage
    public static void main(String[] args) throws Exception {
        File inputDocx = new File("input.docx");
        File outputPdf = new File("output.pdf");

        convertDocxToPdfViaImage(inputDocx, outputPdf);
        System.out.println("PDF generated from DOCX as image.");
    }
}
