/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package infowriter2;

import java.awt.Cursor;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.stage.FileChooser;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import org.apache.commons.compress.archivers.dump.DumpArchiveEntry;
import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.Paragraph;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFHyperlink;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

/**
 *
 * @author HP
 */
public class InfoWriter2 extends JFrame{
    private JButton choose_excel;
    private JButton choose_word;
    private JButton fill;
    private JLabel label_excel, label_word;
    private File file_excel,file_word;
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new InfoWriter2().setVisible(true);
            }
        });
        // TODO code application logic here
    }
    public InfoWriter2(){
        choose_excel = new JButton("Choisir un fichier excel");
        label_excel = new JLabel();
        choose_word = new JButton("choisir un fichier word");
        label_word = new JLabel();
        fill = new JButton("Remplir");
        choose_excel.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){  
                JFileChooser c = new JFileChooser();
                int rVal = c.showOpenDialog(InfoWriter2.this);
                if (rVal == JFileChooser.APPROVE_OPTION) {
                   file_excel = new File("");
                   file_excel = c.getSelectedFile(); 
                   label_excel.setText(file_excel.getName());
                } 
         }});
        choose_word.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){  
                JFileChooser c = new JFileChooser();
                int rVal = c.showOpenDialog(InfoWriter2.this);
                if (rVal == JFileChooser.APPROVE_OPTION) {
                   file_word = new File("");
                   file_word = c.getSelectedFile(); 
                   label_word.setText(file_word.getName());
                } 
         }});
        Path path = Paths.get("C:\\Optimum");
         if (!Files.exists(path)) {
            try {
                Files.createDirectories(path);
            } catch (IOException e) {
                //fail to create directory
                e.printStackTrace();
            }
        }
         fill.addActionListener(new ActionListener(){
            @Override
            public void actionPerformed(ActionEvent e){  
                try {
                    Fill();
                } catch (IOException ex) {
                    Logger.getLogger(InfoWriter2.class.getName()).log(Level.SEVERE, null, ex);
                } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException ex) {
                    Logger.getLogger(InfoWriter2.class.getName()).log(Level.SEVERE, null, ex);
                }
         }});
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints(0, 0, 1, 1, 1.0, 1.0,
            GridBagConstraints.CENTER, GridBagConstraints.NONE, new Insets(
                  50, 10, 0, 0), 0, 0);
        GridBagConstraints gbc2 = new GridBagConstraints(0, 0, 1, 1, 1.0, 1.0,
            GridBagConstraints.CENTER, GridBagConstraints.NONE, new Insets(
                  0, 0, 0, 0), 0, 0);
        panel.add(choose_excel,gbc);
        gbc.gridx = 0;
        gbc.gridy = 1;
        //panel.add(choose_word,gbc);
        gbc.gridx = 1;
        gbc.gridy = 0;
        panel.add(label_excel,gbc);
        gbc.gridx = 1;
        gbc.gridy = 1;
        //panel.add(label_word,gbc);
        gbc.gridx = 0;
        gbc.gridy = 2;
        panel.add(fill,gbc);
        this.setLayout(new GridBagLayout());
        this.getContentPane().add(panel,gbc2);
        panel.setOpaque(false);
        this.setSize(500, 500);
        this.setResizable(false);
        this.setFocusable(true);
        this.setVisible(true);
        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
   }
    
    public void Fill() throws IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException{
       Workbook workbook = null;
         if (file_excel == null){
            JOptionPane.showMessageDialog(null, "Veuillez sélectionner un fichier au format excel");
        }
        else{
            workbook = WorkbookFactory.create(file_excel);
         }
        Sheet firstSheet = workbook.getSheetAt(0);
        Row row;
       /* XWPFDocument document = new XWPFDocument();

      //Write the Document in file system
      FileOutputStream out = null;
        if (file_word == null){
            JOptionPane.showMessageDialog(null, "Veuillez sélectionner un fichier au format word");
        }
        else{
            out = new FileOutputStream(file_word);
        }*/
      //create Paragraph
      XWPFParagraph paragraph ;
      XWPFRun run;
      
      for (int i = 1; i <= firstSheet.getLastRowNum (); i++) {
          
            row=(Row) firstSheet.getRow(i);
            /*if (policy.getDefaultHeader() == null) {
   // Need to create some new headers
   // The easy way, gives a single empty paragraph
   XWPFHeader headerD = policy.createHeader(policy.DEFAULT);
   headerD.addPictureData(new FileInputStream("image.png"), NORMAL);
            }*/            
            //System.out.println(row.getCell(0).getStringCellValue());
            String surname = row.getCell(0).getStringCellValue();
            String name = row.getCell(1).getStringCellValue();
            String ID = row.getCell(4).getStringCellValue();
            File f = new File("C:\\Optimum\\contrat"+" "+ID+" "+surname+" "+name+".docx");
            f.createNewFile();
            XWPFDocument document = new XWPFDocument();
            CTDocument1 document1 = document.getDocument();
            CTBody body = document1.getBody();
            if (!body.isSetSectPr()) {
                 body.addNewSectPr();
            }
            CTSectPr section = body.getSectPr();
            if(!section.isSetPgSz()) {
                section.addNewPgSz();       
            }
            CTPageSz pageSize = section.getPgSz();
            pageSize.setW(BigInteger.valueOf(12240));
            pageSize.setH(BigInteger.valueOf(20160));
            FileOutputStream out = null;
            out = new FileOutputStream(f);
             paragraph = document.createParagraph();
             run = paragraph.createRun();
             //run.addTab();
             run.setText("                   ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setUnderline(UnderlinePatterns.SINGLE);
             run.setText("CONTRAT DE TRAVAIL À DURÉE DÉTERMINÉE À TERME IMPRÉCIS");
             run.addBreak();
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("          I-          ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setUnderline(UnderlinePatterns.SINGLE);
             run.setText("IDENTIFICATION DES PARTIES");
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Entre les soussignés");
             run.addBreak();
             run.addBreak();
             run.setText("OPTIMUM INTERNATIONAL");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(", SARL, au capital de 1 000 000 FCFA, dont le siège est sis ");
             run.addBreak();
             run.setText("Abidjan, Cocody Opération Latrille II plateaux, Aghien Las Palmas, 01 BP 5755 Abidjan 01, ");
             run.addBreak();
             run.setText("téléphones  59 11 00 00/22 52 26 45, immatriculée au Régistre du Commerce et du Crédit Mobilier ");
             run.setText("sous le numéro ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("CI- ABJ-2014-B-21785");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(", représentée par ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Monsieur TRAORE Oumar");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(", en sa ");
             run.setText("qualité de ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("Co-gérant ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText(";");
             run.addBreak();
             run.addBreak();
             run.setText("Ci-après désignée « l’Employeur »                                                                             d’une part ;");
             run.addBreak();
             run.addBreak();
             run.setText("ET ");
             run.addBreak();
             run.addBreak();
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             //run.setBold(false);
             //run.setUnderline(UnderlinePatterns.NONE);
             run.setText("Nom: "+surname);
             run.addBreak();
             run.addBreak();
             run.setText("Prénom(s): "+name);
             run.addBreak();
             run.addBreak();
             String birthday = row.getCell(2).getStringCellValue();
             run.setText("Date et lieu de naissance: "+birthday);
             run.addBreak();
             run.addBreak();
             String nationality = row.getCell(3).getStringCellValue();
             run.setText("Nationalité: "+nationality);
             run.addBreak();
             run.addBreak();
             run.setText("Référence pièce d'identité: "+ID);
             run.addBreak();
             run.addBreak();
             String home = row.getCell(5).getStringCellValue();
             run.setText("Domicile: "+home);
             run.addBreak();
             run.addBreak();
             String representative = row.getCell(6).getStringCellValue();
             run.setText("représentant(e) légal(e): "+representative);
             run.addBreak();
             run.addBreak();
             String representative_phone = row.getCell(7).getStringCellValue();
             run.setText("Contacts représentant(e) légal(e): "+representative_phone);
             run.addBreak();
             run.addBreak();
             String email = row.getCell(8).getStringCellValue();
             run.setText("Adresse email(Obligatoire): "+email);
             run.addBreak();
             run.addBreak();
             String marital_situation = row.getCell(9).getStringCellValue();
             run.setText("Situation matrimoniale: "+marital_situation);
             run.addBreak();
             run.addBreak();
             String phone = row.getCell(10).getStringCellValue();
             run.setText("Contacts téléphoniques: "+phone);
             run.addBreak();
             run.addBreak();
             run.setText("Ci- après désigné « ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setBold(true);
             run.setText("L’employé(e)");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(12);
             run.setText("»                                                                              d’autre part ;");
             run.addBreak();
             run.addBreak();
             //run.addBreak();
             run.setText("Il a été convenu ce qui suit :");
             String function = row.getCell(11).getStringCellValue();
             //run.addBreak();
             //run.addBreak();
             //run.addBreak();
             //run.addBreak();
            //run.addBreak(BreakType.PAGE);
            //run.addBreak(BreakType.);
              CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
             XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, sectPr);
             XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
             paragraph = header.createParagraph();
             run = paragraph.createRun();
             URL header_img1=getClass().getResource("header1.png");
             run.addPicture(header_img1.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header1.png", Units.toEMU(500), Units.toEMU(15));
             paragraph = header.createParagraph();
             XmlCursor cursor = paragraph.getCTP().newCursor();
             //run.setTextPosition(4);
             int twipsPerInch = 1440;
  //create table
  XWPFTable table = header.insertNewTbl(cursor);
  table.setWidth(6*twipsPerInch);
  //create CTTblGrid for this table with widths of the 2 columns. 
  //necessary for Libreoffice/Openoffice to accept the column widths.
  //first column = 2 inches width
  table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(2*twipsPerInch));
  //second column = 4 inches width
  table.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(4*twipsPerInch));
  //create first row
  XWPFTableRow tableRow = table.createRow();
  tableRow.setHeight(1);
  //first cell
  XWPFTableCell cell = tableRow.createCell();
  //set width for first column = 2 inches
  CTTblWidth tblWidth = cell.getCTTc().addNewTcPr().addNewTcW();
  tblWidth.setW(BigInteger.valueOf(2*twipsPerInch));
  //STTblWidth.DXA is used to specify width in twentieths of a point.
  tblWidth.setType(STTblWidth.DXA);
  paragraph = cell.getParagraphArray(0); if (paragraph == null) paragraph = cell.addParagraph();
  //first run in paragraph having picture
  run = paragraph.createRun();
             run.setText("       ");
             URL header_img2=getClass().getResource("logo.png");
             run.addPicture(header_img2.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "logo.png", Units.toEMU(150), Units.toEMU(50));
             cell = tableRow.addNewTableCell();
             CTTblWidth tblWidth2 = cell.getCTTc().addNewTcPr().addNewTcW();
             tblWidth2.setW(BigInteger.valueOf(3*twipsPerInch));
             paragraph = cell.getParagraphArray(0); 
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(8);
             run.setBold(true);
             run.setText("République de Côte d’Ivoire");
             run.addBreak();             
             run.setText("Abidjan, Opération Latrille II Plateaux Cocody Aghien LAS PALMAS");
             run.addBreak();
             run.setText("01 BP 5755 Abidjan 01");
             run.addBreak();
             run.setText("Fixe : + 225 22 52 26 45");
             run.addBreak();
             run.setText("Cel : + 225 02 72 53 53 / + 225 77 92 89 39/");
             run.addBreak();
             run.setText("+ 225 55  09  99  57");
             run.addBreak();
             run.setText("E-mail : ");
             run = paragraph.createRun();
             run.setFontFamily("Times New Roman");
             run.setFontSize(8);
             run.setBold(true);
             run.setColor("0000FF");
             run.setUnderline(UnderlinePatterns.SINGLE);
             run.setText("info@optimum-international.net");
             cell = tableRow.addNewTableCell();
             paragraph = cell.getParagraphArray(0); 
             run = paragraph.createRun();
             run.addBreak();
             run.addBreak();
             run.addBreak();
             run.addBreak();
             run.addBreak();
             run.setFontFamily("Times New Roman");
             run.setFontSize(9);
             run.setBold(true);
             run.setText("RENOUVELLEMNT");
             run.addBreak();
             run.setText(function);
             paragraph = header.createParagraph();
             run = paragraph.createRun();
             URL header_img3=getClass().getResource("header2.png");
             run.addPicture(header_img3.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "header2.png", Units.toEMU(500), Units.toEMU(15));
             /*XWPFHyperlinkRun hyperlinkrun = createHyperlinkRun(paragraph, "info@optimum-international.net");
             hyperlinkrun.setText("info@optimum-international.net");
             hyperlinkrun.setColor();
             hyperlinkrun.setUnderline(UnderlinePatterns.SINGLE);*/
             XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
             paragraph = footer.createParagraph();
             run = paragraph.createRun();
             URL footer_img1=getClass().getResource("footer1.png");
             run.addPicture(footer_img1.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer1.png", Units.toEMU(500), Units.toEMU(5));
             paragraph = footer.createParagraph();
             cursor = paragraph.getCTP().newCursor();
             table = footer.insertNewTbl(cursor);
  table.setWidth(6*twipsPerInch);
  //create CTTblGrid for this table with widths of the 2 columns. 
  //necessary for Libreoffice/Openoffice to accept the column widths.
  //first column = 2 inches width
  table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(2*twipsPerInch));
  //second column = 4 inches width
  table.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(4*twipsPerInch));
  //create first row
  tableRow = table.createRow();
  tableRow.setHeight(1);
  //first cell
  cell = tableRow.createCell();
  //set width for first column = 2 inches
  tblWidth = cell.getCTTc().addNewTcPr().addNewTcW();
  tblWidth.setW(BigInteger.valueOf(8*twipsPerInch));
  //STTblWidth.DXA is used to specify width in twentieths of a point.
  tblWidth.setType(STTblWidth.DXA);
  paragraph = cell.getParagraphArray(0); if (paragraph == null) paragraph = cell.addParagraph();
  run = paragraph.createRun();
  //first run in paragraph having picture
  run = paragraph.createRun();
            run.setFontFamily("Times New Roman");
            run.setFontSize(7);
            run.setBold(true);
            run.setText("     République de Côte d’Ivoire Abidjan, Opération Latrille II Plateaux Cocody Aghien LAS PALMAS 01 BP 5755 Abidjan 01");
            run.addBreak();
            run.setText("  Fixe : + 225 22 52 26 45  Cél : + 225 02 72 53 53 / + 225 77 92 89 39 / + 225 55  09  99  57");
            run.addBreak();
            run.setText("  E-mail : info@optimum-international.net");
            run.addBreak();
            run.setText("        RCCM n  : CI-ABJ-2014-B-21785 - Banque : CI 042 01265 065360000441 77 BIAO CI Danga");
            cell = tableRow.addNewTableCell();
            paragraph = cell.getParagraphArray(0); 
            run = paragraph.createRun();
            URL footer_img2=getClass().getResource("footer2.png");
            run.addPicture(footer_img2.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "footer2.png", Units.toEMU(50), Units.toEMU(20));
            paragraph = footer.createParagraph();
            run = paragraph.createRun();
            run.setFontFamily("Times New Roman");
            run.setFontSize(10);
            run.setText(function+"                                                                                                                     ");
            run = paragraph.createRun();
            run.setFontFamily("Calibri");
            run.setFontSize(11);
            run.setItalic(true);
            run.setText("Version J COR 2019");
            document.write(out);
            out.close();
      }
             /*run.addBreak();
             URL logo=getClass().getResource("logo.png");
             run.addPicture(logo.openStream(), XWPFDocument.PICTURE_TYPE_PNG, "logo.png", Units.toEMU(50), Units.toEMU(50));*/
            //document.write(out);
             //out.close();
     
      JOptionPane.showMessageDialog(null, "Remplissage terminé"); 
    }
     static XWPFHyperlinkRun createHyperlinkRun(XWPFParagraph paragraph, String uri) {
  String rId = paragraph.getDocument().getPackagePart().addExternalRelationship(
    uri, 
    XWPFRelation.HYPERLINK.getRelation()
   ).getId();

  CTHyperlink cthyperLink=paragraph.getCTP().addNewHyperlink();
  cthyperLink.setId(rId);
  cthyperLink.addNewR();

  return new XWPFHyperlinkRun(
    cthyperLink,
    cthyperLink.getRArray(0),
    paragraph
   );
 }   
    }
