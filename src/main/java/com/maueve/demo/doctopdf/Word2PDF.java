
package com.maueve.demo.doctopdf;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;


import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;


//import org.docx4j.Docx4J;

/*
import org.apache.log4j.Level;
import org.docx4j.convert.out.pdf.PdfConversion;
import org.docx4j.convert.out.pdf.viaXSLFO.Conversion;
import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
 */

//import org.docx4j.utils.Log4jConfigurator;

//https://github.com/badalb/doc4j-example/blob/master/doc4j-example/src/main/java/com/testdoctopdf/DocToPDF.java
//https://stackoverrun.com/es/q/11318821
//https://www.docx4java.org/blog/2011/10/hello-maven-central/
//https://stackoverrun.com/es/q/1548450
//https://stackoverflow.com/questions/6201736/javausing-apache-poi-how-to-convert-ms-word-file-to-pdf
//https://stackoverflow.com/questions/51330192/trying-to-make-simple-pdf-document-with-apache-poi


public class Word2PDF {

    public static void main(String[] args) throws IOException, Docx4JException {

        String basePath = "src/main/resources/";
        String archivo = "Test_out.docx";

        String inputPath = basePath + archivo;
        String opPath = basePath;

        Word2PDF w2F = new Word2PDF();

        File src = new File(inputPath);
        File dest = new File(opPath);
        if (!dest.exists()) {
            dest.mkdir();
        }
        if (src.isDirectory()) {
            for (File currentFile : src.listFiles()) {
                w2F.createPDF(currentFile.getAbsolutePath(), opPath);
            }
        } else {
            w2F.createPDF(src.getAbsolutePath(), opPath);
        }
        System.out.println("Done!");
    }

    public void createPDF(final String inputFile, final String opPath) throws IOException, Docx4JException {
        InputStream is = null;
        try {
            // Load doc into WordprocessingMLPackage
            File wordFile = new File(inputFile);
            is = new FileInputStream(wordFile);

            Docx4jProperties.getProperties().setProperty("docx4j.Log4j.Configurator.disabled", "true");
            //Log4jConfigurator.configure();
            //org.docx4j.convert.out.pdf.viaXSLFO.Conversion.log.setLevel(Level.OFF);

            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(is);
            // Prepare Pdf settings
            //PdfSettings pdfSettings = new PdfSettings();
            // Convert WordprocessingMLPackage to Pdf
            OutputStream out = new FileOutputStream(new File(opPath + (wordFile.getName().split("\\.")[0] + ".pdf")));
            //PdfConversion converter = new Conversion(wordMLPackage);
            //converter.output(out, pdfSettings);

            Docx4J.toPDF(wordMLPackage, out);

        } finally {
            if (null != is) {
                is.close();
                is = null;
            }
        }
    }

}
