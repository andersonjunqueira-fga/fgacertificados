import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;

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

public class App {

	private static final String IMAGEPATH = "D:\\desenv\\pessoal\\fga-certificados\\src\\main\\resources\\";
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

	public static void criarCertificados(String certificado, int linhaInicial, int linhaFinal) throws Exception {

		InputStream inp = new FileInputStream(IMAGEPATH + PLANILHA);
		XSSFWorkbook wb = (XSSFWorkbook)WorkbookFactory.create(inp);
		Sheet sheet = wb.getSheetAt(0);

		for(int i = linhaInicial; i <= linhaFinal; i++) {
			System.out.print(i);

			Row row = sheet.getRow(i);
			Cell cell = row.getCell(1);

			String nome = cell.getStringCellValue();
			String pdf = createPDF(nome, certificado, i + "");
			enviar(nome, "anderson.junqueira@gmail.com", pdf);

			System.out.println(" - " + nome + " CRIADO!");
		}

		inp.close();

	}

	public static void enviar(String nome, String email, String filename) throws Exception {

		HtmlEmail msg = new HtmlEmail();
		msg.setHostName("mail.unb.br");
		msg.setSmtpPort(587);
		msg.setAuthenticator(new DefaultAuthenticator("ectgama@unb.br", "ectgama2010"));
		msg.setSSLOnConnect(true);
		msg.setFrom("ectgama@unb.br");
		msg.setSubject("II WORKSHOP ENG. AUTOMOTIVA - Certificado");
		String cid = msg.embed(new File(IMAGEPATH + "logo-email.png"), "logo");

		msg.addTo(email);
		msg.setMsg("<p><img src='cid:" + cid + "' width=\"50%\"></p><p>Prezado(a) {0}</p><p>Voc\\u00EA est\\u00E1 recebendo o seu certificado do II WORKSHOP DE ENGENHARIA AUTOMOTIVA / UnB GAMA.</p><br/><p>Comiss\\u00C3o Organizadora</p>");

		EmailAttachment attachment = new EmailAttachment();
	    attachment.setPath(filename);
	    attachment.setDisposition(EmailAttachment.ATTACHMENT);
	    attachment.setName("Certificado");
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
//		content.moveTextPositionByAmount(270, 280);
		content.drawString(nome);
		content.endText();
		content.close();

		String file = "d:\\desenv\\temp\\" + filename + ".pdf";
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
			doc.save("d:\\desenv\\" + c + ".pdf");
			doc.close();

			System.out.println("SAVED " + c);
		}

		System.out.println("FINISHED.");

	}

	public static PDImageXObject getPDImage(PDDocument doc, String imagepath) throws IOException {
		return PDImageXObject.createFromFile(imagepath, doc);
	}


}
