package org.lxtuong.test.docx4j;

import static org.junit.Assert.*;

import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * 
 * @author Xuan Tuong LE
 *
 */
public class Docx4jBasicTest {
  private static final String DOCX4J = "docx4j";
  private File docx4jTempDirectory;
  
  @Before
  public void setUp() {
    docx4jTempDirectory = new File(FileUtils.getTempDirectoryPath(), DOCX4J);
    assertTrue(docx4jTempDirectory.mkdirs());
  }

  @After
  public void tearDown() {
    try {
      FileUtils.deleteDirectory(docx4jTempDirectory);
    } catch (IOException e) {
      e.printStackTrace();
      fail();
    }
  } 
  
  @Test
  public void testCreateADocx() {
    WordprocessingMLPackage wordMLPackage = null;
    try {
      wordMLPackage = WordprocessingMLPackage.createPackage();
    } catch (InvalidFormatException e) {
      fail(e.getMessage());
    }

    wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Title", "docx4j");

    wordMLPackage.getMainDocumentPart().addParagraphOfText("Test create a docx with docx4j");

    try {
      File docx = new File(docx4jTempDirectory, "testCreateADocx.docx");
      wordMLPackage.save(docx);
    } catch (Docx4JException e) {
      e.printStackTrace();
    }

  }
}
