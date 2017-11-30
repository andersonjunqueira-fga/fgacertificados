import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.commons.mail.DefaultAuthenticator;
import org.apache.commons.mail.EmailAttachment;
import org.apache.commons.mail.HtmlEmail;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDTrueTypeFont;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;

public class App {

    private static final String IMAGEPATH = "C:\\Desenv\\projetos\\fga\\fgacertificados\\src\\main\\resources\\";
    private static final String EVENTO = "evento";
    private static final String COMISSAO = "comissao";
    private static final String EXPOSITOR = "expositor";
    private static final String PALESTRANTE = "palestrante";
    private static final String MINICURSO_MATLAB = "minicurso_matlab";
    private static final String MINICURSO_TECHDAY = "minicurso_techday";

    private static final String PLANILHA = "Certificado_Workshop.xlsx";

    public static void main(String[] args) {
        try {

            Date inicio = new Date(System.currentTimeMillis());
            System.out.println("INICIADO EM : " + inicio);

            criarCertificados();

            Date fim = new Date(System.currentTimeMillis());
            System.out.println("FINALIZADO EM : " + fim);
            System.out.println("TEMPO : " + ((fim.getTime() - inicio.getTime()) / 1000) );

        } catch(Exception ex) {
            ex.printStackTrace();
        }
    }

    public static void criarCertificados() throws Exception {

        // PARTICIPAÇÃO - 8 - 71
        criarCertificados(EVENTO, 8, 71);

        // MINICURSO TECHDAY - 77 - 86
        criarCertificados(MINICURSO_TECHDAY, 77, 86);

        // MATLAB - 91 - 107
        criarCertificados(MINICURSO_MATLAB, 91, 107);

        // EXPOSITOR - 113 - 119
        criarCertificados(EXPOSITOR, 113, 119);

        // COMISSAO DISCENTES - 125 - 133
        criarCertificados(COMISSAO, 125, 133);

        // COMISSAO DOCENTES - 137 - 142
        criarCertificados(COMISSAO, 137, 142);

        // PALESTRANTES - 148 - 156
        criarCertificados(PALESTRANTE, 148, 156);

    }

    public static List<Usuario> loadUsuarios() throws Exception {
        FileReader reader = new FileReader(new File(IMAGEPATH + "usuarios-20170809.json"));
        Type listType = new TypeToken<ArrayList<Usuario>>(){}.getType();
        List<Usuario> yourClassList = new Gson().fromJson(reader, listType);
        return yourClassList;
    }

    public static void criarCertificados(String certificado, int linhaInicial, int linhaFinal) throws Exception {

        List<Usuario> usuarios = loadUsuarios();

        InputStream inp = new FileInputStream(IMAGEPATH + PLANILHA);
        XSSFWorkbook wb = (XSSFWorkbook)WorkbookFactory.create(inp);
        Sheet sheet = wb.getSheetAt(0);

        for(int i = linhaInicial; i <= linhaFinal; i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(1);

            String nome = cell.getStringCellValue();
            String pdf = createPDF(nome, certificado, i + "");

            List<Object> u = usuarios.stream()
                .filter(item -> item.getNome().equalsIgnoreCase(nome))
                .collect(Collectors.toList());

            if(u.size() > 0) {
//                enviar(nome, "suzana.avila@gmail.com", pdf);
//                System.exit(0);
                enviar(nome, ((Usuario)u.get(0)).getEmail(), pdf);
                System.out.println(" - " + nome + " ENVIADO!");
            } else {
                //System.out.println(nome);
            }

        }

        inp.close();

    }

    public static void enviar(String nome, String email, String filename) throws Exception {

        HtmlEmail msg = new HtmlEmail();
        msg.setHostName("mail.unb.br");
        msg.setAuthenticator(new DefaultAuthenticator("avilas@unb.br", "bsb2005"));
        msg.setSmtpPort(587);

        msg.setFrom("avilas@unb.br");
        msg.setSubject("II WORKSHOP ENG. AUTOMOTIVA - Certificado");
        String cid = msg.embed(new File(IMAGEPATH + "logo-email.png"), "logo");

        msg.addTo(email);
        msg.setMsg("<p><img src='cid:" + cid + "' width=\"50%\"></p><p>Prezado(a),</p><p>Você está recebendo o seu certificado do II WORKSHOP DE ENGENHARIA AUTOMOTIVA / UnB GAMA.</p><br/><p>Comissão Organizadora</p>");

        EmailAttachment attachment = new EmailAttachment();
        attachment.setPath(filename);
        attachment.setDisposition(EmailAttachment.ATTACHMENT);
        attachment.setName("certificado.pdf");
        msg.attach(attachment);

        msg.send();

    }


    public static String createPDF(String nome, String tipo, String filename) throws Exception {

        int fsize = 24;

        PDDocument doc = PDDocument.load(new File(IMAGEPATH + tipo + ".pdf"));
        PDFont font = PDTrueTypeFont.loadTTF(doc, new File(IMAGEPATH + "MTCORSVA.TTF"));
        PDPage firstPage = doc.getPage(0);

        PDPageContentStream content = new PDPageContentStream(doc, firstPage, true, true);

        float nomeWidth = font.getStringWidth(nome) / 1000 * fsize;
        float nomeHeight = font.getFontDescriptor().getFontBoundingBox().getHeight() / 1000 * fsize;

        content.setNonStrokingColor(0, 0, 0);
        content.beginText();
        content.setFont(font, fsize);
        content.moveTextPositionByAmount((firstPage.getMediaBox().getWidth() - nomeWidth) / 2, 280);
        content.drawString(nome);
        content.endText();
        content.close();

        String file = "c:\\desenv\\temp\\" + filename + ".pdf";
        doc.save(file);
        doc.close();

        return file;

    }

    public static void createMatrizes() throws Exception {

        String[] certs = new String[] { EVENTO, COMISSAO, EXPOSITOR, PALESTRANTE, MINICURSO_MATLAB, MINICURSO_TECHDAY };

        System.out.println("STARTING ...");

        for(String c : certs) {
            System.out.println("GENERATING " + c);

            PDDocument doc = new PDDocument();
            PDImageXObject pdImage = getPDImage(doc, IMAGEPATH + c + ".png");

            PDPage page = new PDPage(new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth()));
            PDPageContentStream contents = new PDPageContentStream(doc, page);
            contents.drawImage(pdImage, 0, 0, PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth());
            contents.close();

            doc.addPage(page);
            doc.save("c:\\desenv\\temp\\" + c + ".pdf");
            doc.close();

            System.out.println("SAVED " + c);
        }

        System.out.println("FINISHED.");

    }

    public static PDImageXObject getPDImage(PDDocument doc, String imagepath) throws IOException {
        return PDImageXObject.createFromFile(imagepath, doc);
    }


}
