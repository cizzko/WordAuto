import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

public class Main {
    static String fontName = "Times New Roman";
    static File outDir, inDir;
    static int fontSize = 14;
    static double spacing = 1.5;

    public static void main(String[] args) {
        readConf();

        //для полноценной работы внешних вба требуется вручную разрешить это в ворде
        //для обхода создан пустой файл с готовым скриптом
        //в пустышку полностью копируется нужный файл и запускается скрипт
        //после завершения результат копируется обратно
        File templateFile = new File("src/assets/zaglushka.docm");

        ActiveXComponent word = new ActiveXComponent("Word.Application");
        word.setProperty("Visible", new Variant(false));

        try {
            Dispatch documents = word.getProperty("Documents").toDispatch();
            File[] files = inDir.listFiles(f -> f.isFile() && f.getName().endsWith(".docx"));

            for (File targetDoc : files) {
                String outputName = targetDoc.getName();
                File outputFile = new File(outDir, outputName);

                Dispatch doc = Dispatch.call(documents, "Open", templateFile.getAbsolutePath()).toDispatch();
                Dispatch selection = word.getProperty("Selection").toDispatch();

                Dispatch.call(selection, "InsertFile", targetDoc.getAbsolutePath());
                Dispatch.call(word, "Run", "UniversalFormat", new Variant(fontName), new Variant(fontSize), new Variant(spacing));
                Dispatch.call(doc, "SaveAs", outputFile.getAbsolutePath(), new Variant(16));
                Dispatch.call(doc, "Close");

                //место для логгирования, но не входило в мое тз
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            word.invoke("Quit");
        }
    }

    private static void readConf() {
        try {
            List<String> lines = Files.readAllLines(Path.of("src/assets/config.txt"));

            if (!lines.isEmpty()) {
                String[] parts = lines.getFirst().split(";");

                fontName = parts[0];
                fontSize = Integer.parseInt(parts[1]);
                spacing = Double.parseDouble(parts[2]);
                outDir = new File(parts[3]);
                inDir = new File(parts[4]);
            }
        } catch (Exception e) {
            System.err.println(e);
        }
        //создает каталог если нету, необязателен, но приятно
        outDir.mkdirs();
    }
}