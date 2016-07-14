/*********************************************************************
 * CLASS NAME: ReportUtil                                        
 * © Copyright IBM Corporation 2016. All rights reserved.            
 * 
 * CHANGE ACTIVITY:
 * DATE     | AUTHOR     | ACTIVITY                    
 *----------+------------+-------------------------------------------
 * 04/01/16 | TCOE       | Initial Version          
 *----------+------------+-------------------------------------------
 *          |            |                                           
 ********************************************************************/
package util;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Vector;

import com.lowagie.text.Chunk;
import com.lowagie.text.Document;
import com.lowagie.text.HeaderFooter;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.rtf.RtfWriter2;
import com.lowagie.text.rtf.headerfooter.RtfHeaderFooter;
import com.lowagie.text.rtf.style.RtfFont;

public class ReportUtil {

	private String reportFileName, reportFileFoler;

	private String reportFileFullPath;

	private String fontString = "Courier New";

	// private boolean isFirst = true;

	RtfFont objectFont = new RtfFont(fontString, 12, RtfFont.STYLE_UNDERLINE);
	RtfFont titleFont = new RtfFont(fontString, 12, RtfFont.BOLD);
	RtfFont normalFont = new RtfFont(fontString, 12);
	RtfFont sqlFont = new RtfFont(fontString, 9);
	RtfFont headerFont = new RtfFont(fontString, 8);

	Document doc = null;
	FileOutputStream steam = null;
	RtfWriter2 rtfWriter2 = null;

	public ReportUtil(String testCaseReportName, String reportFileFoler) {
		reportFileName = testCaseReportName;
		this.reportFileFoler = reportFileFoler;
		// setReportFileFullPath();
		setReportFileFullPathWithoutTime();
		setReportFile();
	}

	// private void setReportFileFullPath() {
	// reportFileFullPath = reportFileFoler + "/" +
	// reportFileName+DateTimeUtil.getCurrentDateTime()+".rtf";
	// }

	private void setReportFileFullPathWithoutTime() {
		// reportFileFullPath = reportFileFoler + "/" + reportFileName+".rtf";
		reportFileFullPath = reportFileFoler + "/" + reportFileName + ".doc";
	}

	private void setReportFile() {

		try {
			File report_file = new File(reportFileFullPath);

			if (!report_file.exists()) {
				report_file.createNewFile();
			}
			doc = new Document(PageSize.A4);
			steam = new FileOutputStream(reportFileFullPath, true);
			rtfWriter2 = RtfWriter2.getInstance(doc, steam);
			doc.open();
			// rtfWriter2.importRtfDocument(new FileInputStream(report_file));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void printReport(String msg) {

		String curr_date_curr_zone = DateTimeUtil
				.formatedTime("yyyy-MM-dd HH:mm:ss:SSS");

		switch (msg.trim()) {
		case "START":
			try {
				Paragraph p = new Paragraph();
				p.setFont(objectFont);
				Chunk c = new Chunk("Test Case Name :");
				c.setFont(titleFont);
				p.add(c);
				Chunk c1 = new Chunk(reportFileName);
				c1.setFont(titleFont);
				p.add(c1);
				// doc.add(p);

				p = new Paragraph();
				// doc.add(p);

				p = new Paragraph();
				c = new Chunk("Report Date :");
				c.setFont(objectFont);
				p.add(c);
				c1 = new Chunk(curr_date_curr_zone);
				c1.setFont(normalFont);
				p.add(c1);
				// doc.add(p);

				p = new Paragraph();
				c = new Chunk("File Name :");
				c.setFont(objectFont);
				p.add(c);
				c1 = new Chunk(reportFileFullPath);
				c1.setFont(normalFont);
				p.add(c1);
				// doc.add(p);

				p = new Paragraph();
				doc.add(p);

			} catch (Exception e) {
				e.printStackTrace();
			}
			break;
		case "END":
			try {
				Paragraph p = new Paragraph();
				doc.add(p);

				p = new Paragraph();
				Chunk c = new Chunk("End Date :");
				c.setFont(objectFont);
				p.add(c);
				Chunk c1 = new Chunk(curr_date_curr_zone);
				c1.setFont(normalFont);
				p.add(c1);
				// doc.add(p);
				doc.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
			break;
		default:
			try {
				Paragraph p = new Paragraph();
				doc.add(p);

				p = new Paragraph();
				Chunk c = new Chunk(msg);
				c.setFont(normalFont);
				p.add(c);
				doc.add(p);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	public void printImageIntoReport(File imageFile) {

		try {
			Image img = Image.getInstance(imageFile.getAbsolutePath());
			img.scalePercent(30, 30);
			img.setAlignment(Image.MIDDLE);
			Paragraph p = new Paragraph();
			Chunk c = new Chunk(img, 0, 0);
			p.add(c);
			doc.add(p);
			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void printImageIntoReport(String imageFileFullPath) {
		try {
			Image img = Image.getInstance(imageFileFullPath);
			Paragraph p = new Paragraph();
			Chunk c = new Chunk(img, 0, 0);
			p.add(c);
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void printTestObject(String message) {

		try {
			Paragraph p = new Paragraph();
			p.setFont(objectFont);
			Chunk c = new Chunk("Test Objectives:");
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(message);
			c.setFont(normalFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void printStep(String message, int stepNumber) {

		try {
			Paragraph p = new Paragraph();
			Chunk c = new Chunk(" ");
			c.setNewPage();
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			p.setFont(objectFont);
			c = new Chunk("Test Step " + stepNumber);
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			doc.add(p);

			p = new Paragraph();
			p.setFont(objectFont);
			c = new Chunk("Description (Actions)");
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(message);
			c.setFont(normalFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public void printExpectResult(String message) {

		try {
			Paragraph p = new Paragraph();
			Chunk c = new Chunk(" ");
			doc.add(p);

			p = new Paragraph();
			p.setFont(objectFont);
			c = new Chunk("Expected Results:");
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(message);
			c.setFont(normalFont);
			p.add(c);
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void printActualResult(String message) {

		try {
			Paragraph p = new Paragraph();
			Chunk c = new Chunk(" ");
			doc.add(p);

			p = new Paragraph();
			p.setFont(objectFont);
			c = new Chunk("Actual Results:");
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(message);
			c.setFont(normalFont);
			p.add(c);
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void printSQLResult(Vector<String> result) {

		try {
			Paragraph p = new Paragraph();
			Chunk c = new Chunk(" ");
			doc.add(p);

			p = new Paragraph();
			p.setFont(objectFont);
			c = new Chunk("SQL:");
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			for (int i = 0; i < getMultiRowText(result.get(0)).length; i++) {
				p = new Paragraph();
				c = new Chunk(getMultiRowText(result.get(0))[i]);
				c.setFont(sqlFont);
				p.add(c);
				doc.add(p);
			}

			p = new Paragraph();
			p.setFont(objectFont);
			c = new Chunk("SQL Result:");
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			for (int i = 0; i < getMultiRowText(result.get(1)).length; i++) {
				p = new Paragraph();
				c = new Chunk(getMultiRowText(result.get(1))[i].replace("\r",
						""));
				c.setFont(sqlFont);
				p.add(c);
				doc.add(p);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private String[] getMultiRowText(String s) {
		String[] ss = s.split("\n\r");
		return ss;
	}

	public void printTestDescription(String message) {

		try {
			Paragraph p = new Paragraph();
			p.setFont(objectFont);
			Chunk c = new Chunk("Test Description:");
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(message);
			c.setFont(normalFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void printTestPrerequisites(String message) {

		try {
			Paragraph p = new Paragraph();
			p.setFont(objectFont);
			Chunk c = new Chunk("Test Prerequisites:");
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(message);
			c.setFont(normalFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void printTestPreCondition(String message) {

		try {
			Paragraph p = new Paragraph();
			p.setFont(objectFont);
			Chunk c = new Chunk("Pre-Condition:");
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(message);
			c.setFont(normalFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void printTestEnvironment(String website, String brower) {

		try {
			Paragraph p = new Paragraph();
			p.setFont(objectFont);
			Chunk c = new Chunk("Environment:");
			c.setFont(titleFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(website);
			c.setFont(normalFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk("Brower: " + brower);
			c.setFont(normalFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void printMessage(String message) {

		try {
			Paragraph p = new Paragraph();

			p = new Paragraph();
			Chunk c = new Chunk(message);
			c.setFont(normalFont);
			p.add(c);
			doc.add(p);

			p = new Paragraph();
			c = new Chunk(" ");
			p.add(c);
			doc.add(p);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// Add by Bruce
	public void printPageHeader(String header) {
		try {
			Paragraph headerPara1 = new Paragraph(header, headerFont);
			// headerPara1.setAlignment(HeaderFooter.ALIGN_CENTER);

			RtfHeaderFooter header1 = new RtfHeaderFooter(headerPara1);

			Paragraph headerPara2 = new Paragraph();
			// headerPara2.add(headerPara1);
			headerPara2.add(header1);
			headerPara2.setAlignment(HeaderFooter.ALIGN_LEFT);
			doc.add(headerPara2);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}