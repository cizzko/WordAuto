import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

public class Main {
    public static void main(String[] args) {

        //TODO: абс пути -> относительные
        //todo нейминг
        String fontName = "Times New Roman";
        int fontSize = 14;
        double spacing = 1.5;

        try {
            List<String> lines = Files.readAllLines(Paths.get("C:\\other\\parserJaba\\parser\\src\\assets\\config.txt"));
            if (!lines.isEmpty()) {
                String[] parts = lines.get(0).split(";");

                fontName = parts[0];
                fontSize = Integer.parseInt(parts[1]);
                spacing = Double.parseDouble(parts[2]);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        ActiveXComponent word = new ActiveXComponent("Word.Application");
        word.setProperty("Visible", new Variant(true));

        try {
            Dispatch documents = word.getProperty("Documents").toDispatch();

            File templateFile = new File("C:\\other\\parserJaba\\parser\\src\\assets\\zaglushka.docm");

            Dispatch doc = Dispatch.call(documents, "Open", templateFile.getAbsolutePath()).toDispatch();
            Dispatch select = word.getProperty("Selection").toDispatch();

            Dispatch.call(select, "InsertFile", new File("C:\\other\\parserAssets\\Документ Microsoft Word.docx").getAbsolutePath());
            Dispatch.call(word, "Run", "UniversalFormat", new Variant(fontName), new Variant(fontSize), new Variant(spacing));
            Dispatch.call(doc, "SaveAs", "C:\\other\\parserAssets\\result.docx");

            //todo чтоб сохраняться успевало надумать покрасивее
            Thread.sleep(5000);

            Dispatch.call(doc, "Close", new Variant(false));

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            word.invoke("Quit");
        }
    }
}