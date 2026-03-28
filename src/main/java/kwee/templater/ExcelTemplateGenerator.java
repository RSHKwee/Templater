package kwee.templater;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.VelocityEngine;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

public class ExcelTemplateGenerator {

  private final VelocityEngine velocityEngine;
  private final String templatePath;

  // Constructor met template path
  public ExcelTemplateGenerator(String templatePath) throws Exception {
    this.templatePath = templatePath;

    Properties props = new Properties();

    // Controleer of template directory bestaat
    File templateDir = new File(templatePath);
    if (!templateDir.exists()) {
      // Probeer alternatieve locaties
      String[] alternatieven = { "./templates", "./src/templates", "./src/main/resources/templates", templatePath };

      boolean found = false;
      for (String alt : alternatieven) {
        File testDir = new File(alt);
        if (testDir.exists() && testDir.isDirectory()) {
          templatePath = alt;
          found = true;
          System.out.println("Template directory gevonden: " + testDir.getAbsolutePath());
          break;
        }
      }

      if (!found) {
        throw new Exception("Template directory niet gevonden. Huidige directory: " + new File(".").getAbsolutePath());
      }
    }

    props.setProperty("resource.loader", "file");
    props.setProperty("file.resource.loader.path", templatePath);
    this.velocityEngine = new VelocityEngine(props);
    this.velocityEngine.init();
  }

  /**
   * Genereer bestanden op basis van Excel en template
   * 
   * @param excelPad        Pad naar Excel bestand
   * @param templateNaam    Naam van template bestand
   * @param outputDirectory Output directory voor gegenereerde bestanden
   * @param outputPrefix    Prefix voor gegenereerde bestanden
   * @throws Exception
   */
  public void genereerBestanden(String excelPad, String templateNaam, String outputDirectory, String outputPrefix)
      throws Exception {
    // Lees Excel bestand
    List<Map<String, String>> excelData = leesExcelBestand(excelPad);

    // Laad template
    Template template = velocityEngine.getTemplate(templateNaam, "UTF-8");

    // Genereer bestand voor elke rij
    int rijNummer = 1;
    for (Map<String, String> rijData : excelData) {
      String outputBestand = outputDirectory + File.separator + outputPrefix + "_" + rijNummer + ".txt";

      // Maak context met data uit de rij
      VelocityContext context = new VelocityContext();
      for (Map.Entry<String, String> entry : rijData.entrySet()) {
        context.put(entry.getKey(), entry.getValue());
      }

      // Voeg eventueel extra metadata toe
      context.put("rijNummer", String.valueOf(rijNummer));
      context.put("huidigeDatum", new Date());

      // Genereer output
      genereerBestandVanTemplate(template, context, outputBestand);

      System.out.println("Gegenereerd: " + outputBestand);
      rijNummer++;
    }

    System.out.println("Klaar! " + (rijNummer - 1) + " bestanden gegenereerd.");
  }

  /**
   * Lees Excel bestand en converteer naar lijst van maps
   */
  private List<Map<String, String>> leesExcelBestand(String excelPad) throws IOException {
    List<Map<String, String>> resultaat = new ArrayList<>();

    try (FileInputStream fis = new FileInputStream(excelPad); Workbook workbook = new XSSFWorkbook(fis)) {

      Sheet sheet = workbook.getSheetAt(0);
      Row headerRow = sheet.getRow(0);

      if (headerRow == null) {
        throw new IllegalArgumentException("Excel bestand heeft geen headers");
      }

      // Lees headers
      List<String> headers = new ArrayList<>();
      for (Cell cell : headerRow) {
        headers.add(getCellWaardeAlsString(cell));
      }

      // Verwerk elke data rij
      for (int i = 1; i <= sheet.getLastRowNum(); i++) {
        Row rij = sheet.getRow(i);
        if (rij == null)
          continue;

        Map<String, String> rijData = new HashMap<>();
        for (int j = 0; j < headers.size(); j++) {
          Cell cel = rij.getCell(j);
          String waarde = cel != null ? getCellWaardeAlsString(cel) : "";
          rijData.put(headers.get(j), waarde);
        }

        resultaat.add(rijData);
      }
    }

    return resultaat;
  }

  /**
   * Haal celwaarde op als string
   */
  private String getCellWaardeAlsString(Cell cel) {
    if (cel == null)
      return "";

    switch (cel.getCellType()) {
    case STRING:
      return cel.getStringCellValue();
    case NUMERIC:
      if (DateUtil.isCellDateFormatted(cel)) {
        return cel.getDateCellValue().toString();
      } else {
        double getal = cel.getNumericCellValue();
        if (getal == (long) getal) {
          return String.valueOf((long) getal);
        } else {
          return String.valueOf(getal);
        }
      }
    case BOOLEAN:
      return String.valueOf(cel.getBooleanCellValue());
    case FORMULA:
      try {
        return String.valueOf(cel.getNumericCellValue());
      } catch (Exception e) {
        return cel.getCellFormula();
      }
    default:
      return "";
    }
  }

  /**
   * Genereer bestand op basis van template
   */
  private void genereerBestandVanTemplate(Template template, VelocityContext context, String outputPad)
      throws Exception {
    // Zorg dat output directory bestaat
    Files.createDirectories(Paths.get(outputPad).getParent());

    try (Writer writer = new OutputStreamWriter(new FileOutputStream(outputPad), "UTF-8")) {
      template.merge(context, writer);
    }
  }

  /**
   * Main methode met voorbeeld
   */
  public static void main(String[] args) {
    try {
      // Gebruik de constructor met template pad
      ExcelTemplateGenerator generator = new ExcelTemplateGenerator("./templates");

      // Voorbeeld parameters
      String excelPad = "D:\\Dev\\Github\\tools\\Templater\\target\\resources\\Data\\Data.xlsx";
      String templateNaam = "Template_1.vm";
      String outputDirectory = "output";
      String outputPrefix = "document";

      // Roep de methode aan
      generator.genereerBestanden(excelPad, templateNaam, outputDirectory, outputPrefix);

    } catch (Exception e) {
      System.err.println("Fout bij genereren: " + e.getMessage());
      e.printStackTrace();
    }
  }
}