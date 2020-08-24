package com.maueve.demo.doctopdf;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;

public class DocxToPDF {

    public static void main(String[] args) throws Docx4JException, FileNotFoundException {

        String basePath = "src/main/resources/";
        String archivo = "Test_out4.docx";
        String salida = "Test_out.pdf";

        //WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(basePath + archivo));
        WordprocessingMLPackage wordMLPackage = Docx4J.load(new File(basePath + archivo));
        OutputStream out = new FileOutputStream(new File(basePath + salida));
        Docx4J.toPDF(wordMLPackage, out);
    }
}
