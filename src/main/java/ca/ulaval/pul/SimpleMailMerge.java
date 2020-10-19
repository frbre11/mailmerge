package ca.ulaval.pul;

import com.spire.doc.Document;  
import com.spire.doc.FileFormat;  
import com.spire.doc.reporting.MergeImageFieldEventArgs;  
import com.spire.doc.reporting.MergeImageFieldEventHandler;  
  
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
  
public class SimpleMailMerge {  
    public static void main(String[] args) throws Exception {  
        //create a Document instance  
        Document document = new Document();  
        //load the template document  
        document.loadFromFile("template.docx");  
  
        String prenom = "Rollande";
        String nom = "Théberge";
        Date dateTmp = new Date();  
        SimpleDateFormat format = new SimpleDateFormat("dd-MM-yyyy");
        String dateCourante = format.format(dateTmp);
        dateTmp = new GregorianCalendar(1954, Calendar.FEBRUARY, 11).getTime();
        String dateNaissance = format.format(dateTmp);  
        String[] filedNames = new String[]{"date", "titre", "prenomMedecin", "nomMedecin", "prenom", "nom", "dateNaissance", "noCivique", "adresse", "ville", "province", "codePostal", "risque1", "risque2"};  
        String[] filedValues = new String[]{dateCourante, "Dr.", "Herménégilde", "Laliberté", prenom, nom, dateNaissance, "1234", "rue de l'Église", "St-Onésime", "QC", "H0H 0H0", "11", "22"};  
        document.getMailMerge().MergeImageField = new MergeImageFieldEventHandler() {  
            @Override  
            public void invoke(Object sender, MergeImageFieldEventArgs args) {  
                mailMerge_MergeImageField(sender, args);  
            }  
        };  
        //call execute method to merge image and string values to the document  
        document.getMailMerge().execute(filedNames, filedValues);  
  
        //save the document  
        document.saveToFile(prenom + "." + nom + ".pdf", FileFormat.PDF);  
    }  
    private static void mailMerge_MergeImageField(Object sender, MergeImageFieldEventArgs field) {  
        String filePath = field.getImageFileName();  
        if (filePath != null && !"".equals(filePath)) {  
            try {  
                field.setImage(filePath);  
            } catch (Exception e) {  
                e.printStackTrace();  
            }  
        }  
    }  
}  

