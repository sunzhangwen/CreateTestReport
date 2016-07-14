package script;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;

import util.DateTimeUtil;
import util.ExcelUtil;
import util.ReportUtil;

/*
 * CHANGE ACTIVITY:
 1. Initial Version 06/02/16  Sun Zhang Wen
 2. 06/10/2016  Sun Zhang Wen
 2.1 Add additional statement to process "Prerequisites" and "Screen Shot" item
 2.2 Modify the structure
 3. 07/14/2016  sun Zhang Wen
 3.1 Move to Github
 */
public class CreateTestReport {

	/*--------------------------------------
	 * Global Setting
	 * --------------------------------------*/
	final private static String TESTCASE_PATH = "./test case/";
	final private static String TESTREPORT_PATH = "./test report";
	final private static String IMAGE_PATH = "./screen shot/";
	final private static String IMAGE_FORMAT = ".png";
	final private static String FILE_HEADER_1 = ";   My_test;   Tester: ";
	final private static String FILE_HEADER_2 = ";   Date Tested: ";
	/*--------------------------------------
	 * User Setting
	 * --------------------------------------*/
	final private static String SUBJECTNAME = "Phase 1";
	final private static String CASEFILE_NAME = "TC.xls";
	final private static String SHEETNAME = "Demo";
	final private static String USER = "SunZhangWen";

	static ArrayList<HashMap<String, String[]>> CaseList = new ArrayList<HashMap<String, String[]>>();

	public static void main(String[] args) {

		CaseList = ExcelUtil.getTestCase(TESTCASE_PATH + CASEFILE_NAME,
				SHEETNAME, SUBJECTNAME);

		System.out.println("\nTotal test case: " + CaseList.size());

		for (int caseIndex = 0; caseIndex < CaseList.size(); caseIndex++) {
			ReportUtil report = new ReportUtil(CaseList.get(caseIndex).get(
					"Test Name")[0], TESTREPORT_PATH);
			report.printReport("START");
			report.printPageHeader(CaseList.get(caseIndex).get("Test Name")[0]
					+ FILE_HEADER_1 + USER + FILE_HEADER_2
					+ DateTimeUtil.formatedTime("MM-dd-yyyy"));
			report.printTestDescription(CaseList.get(caseIndex).get(
					"Description")[0]);
			if (!CaseList.get(caseIndex).get("Prerequisites")[0].equals("")) {
				report.printTestPrerequisites(CaseList.get(caseIndex).get(
						"Prerequisites")[0]);
			} else {
				report.printTestPrerequisites("N/A");
			}

			report.printTestEnvironment(
					CaseList.get(caseIndex).get("Website")[0],
					CaseList.get(caseIndex).get("Browser")[0]);

			for (int i = 0; i < CaseList.get(caseIndex).get("Step Description").length; i++) {

				report.printStep(CaseList.get(caseIndex)
						.get("Step Description")[i], i + 1);
				report.printExpectResult(CaseList.get(caseIndex).get(
						"Expected Results")[i]);
				if (!CaseList.get(caseIndex).get("Screen Shot")[i].equals("")) {
					if (CaseList.get(caseIndex).get("Screen Shot")[i]
							.contains(";")) {
						String[] image = CaseList.get(caseIndex).get(
								"Screen Shot")[i].split(";");
						for (int j = 0; j < image.length; j++) {
							report.printImageIntoReport(new File(IMAGE_PATH
									+ image[j] + IMAGE_FORMAT));
						}

					} else {
						report.printImageIntoReport(new File(IMAGE_PATH
								+ CaseList.get(caseIndex).get("Screen Shot")[i]
								+ IMAGE_FORMAT));
					}
				}
				report.printActualResult(CaseList.get(caseIndex).get(
						"Actual Results")[i]);
			}

			report.printReport("END");
			System.out.println("\nTest case: "
					+ CaseList.get(caseIndex).get("Test Name")[0]
					+ "\nCreate successfully");
		}
	}
}
