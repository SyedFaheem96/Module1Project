package ExcelMail;

public class Main {

	private static String toCreateFile  = "DoseCheck.xls";
	public static String highlightedFile  = "DoseCheck2.xls";
	
	public static void main(String[] args) throws Exception {
		
		ExcelCreator creator = new ExcelCreator();
		creator.start(toCreateFile, highlightedFile);
		
		MailSender.send();
	}
}