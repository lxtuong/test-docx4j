package org.lxtuong.test.docx4j;

import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.JAXBElement;

import org.apache.commons.io.FileUtils;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.SpreadsheetML.SharedStrings;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.STCellType;
import org.xlsx4j.sml.SheetData;
import org.xlsx4j.sml.Worksheet;

/**
 * 
 * @author Xuan Tuong LE
 * 
 */
public class Docx4jBasicTest {
  private static final String DOCX4J = "docx4j";
  private File docx4jTempDirectory;
  private static List<WorksheetPart> worksheets = new ArrayList<WorksheetPart>();

  private static SharedStrings sharedStrings = null;

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

  private File getFileFromResource(String filename) {
    URL url = this.getClass().getResource("/" + filename);    
    return new File(url.getFile());
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

  @Test
  public void readAXlsx() {

    try {
      SpreadsheetMLPackage xlsxPkg = SpreadsheetMLPackage.load(getFileFromResource("test.xlsx"));
   // List the parts by walking the rels tree
      RelationshipsPart rp = xlsxPkg.getRelationshipsPart();
      StringBuilder sb = new StringBuilder();
      printInfo(rp, sb, "");
      System.out.println(sb.toString());

      // Now lets print the cell content
      for(WorksheetPart sheet: worksheets) {
        System.out.println(sheet.getPartName().getName() );
        Worksheet ws = sheet.getJaxbElement();
        SheetData data = ws.getSheetData();
        for (Row r : data.getRow() ) {
          System.out.println("row " + r.getR() );       
          for (Cell c : r.getC() ) {
            if (c.getT().equals(STCellType.S)) {
              System.out.println( "  " + c.getR() + " contains " +
                  sharedStrings.getJaxbElement().getSi().get(Integer.parseInt(c.getV())).getT()
                      );
            } else {
              // TODO: handle other cell types
              System.out.println( "  " + c.getR() + " contains " + c.getV() );
            }
          }
        }
      }
    } catch (Docx4JException e) {
      e.printStackTrace();
      fail(e.getMessage());
    }
  }
  
  public static void  printInfo(Part p, StringBuilder sb, String indent) {
    sb.append("\n" + indent + "Part " + p.getPartName() + " [" + p.getClass().getName() + "] " );   
    if (p instanceof JaxbXmlPart) {
      Object o = ((JaxbXmlPart)p).getJaxbElement();
      if (o instanceof javax.xml.bind.JAXBElement) {
        sb.append(" containing JaxbElement:" + XmlUtils.JAXBElementDebug((JAXBElement)o) );
      } else {
        sb.append(" containing JaxbElement:"  + o.getClass().getName() );
      }
    }
    if (p instanceof WorksheetPart) {
      worksheets.add((WorksheetPart)p);
    } else if (p instanceof SharedStrings) {
      sharedStrings = (SharedStrings)p;
    }

  }

}
