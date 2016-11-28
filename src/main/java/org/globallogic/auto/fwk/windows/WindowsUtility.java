package org.globallogic.auto.fwk.windows;

import java.io.File;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class WindowsUtility {

	public static boolean writeToExcel(String filename, String content) {
		return writeToExcel(filename, content, "A1");
	}

	public static boolean writeToExcel(String filename, String content, String cellref) {

		boolean result = false;

		try {
			System.out.println("Atempting to write to file:" + filename + "\ncontent=" + content
					+ "\ntarget cell reference=" + cellref);
			File f = new File(System.getProperty("user.dir") + "\\" + filename);
			System.out.println("user dir=" + f.getAbsolutePath());
			String targetFileName = f.getAbsolutePath();

			ActiveXComponent component = new ActiveXComponent("Excel.Application");
			Dispatch.put(component, "Visible", new Variant(true));
			Dispatch.put(component, "DisplayAlerts", new Variant(false));
			// get a handle to current workbook (or default new one)
			Dispatch wkbks = component.getProperty("Workbooks").toDispatch();
			Dispatch wkbksingle;
			if (f.exists()) {
				wkbksingle = Dispatch
						.invoke(wkbks, "Open", Dispatch.Get, new Object[] { targetFileName }, new int[] { 1 })
						.toDispatch();
				// Dispatch.get(wkbks, "Open").toDispatch();
			} else {

				// call the add function to add your own sheet
				wkbksingle = Dispatch.get(wkbks, "Add").toDispatch();
			}
			// now get a handle to the active sheet
			Dispatch sheet = Dispatch.get(wkbksingle, "ActiveSheet").toDispatch();

			Dispatch cell1 = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] { cellref }, new int[] { 1 })
					.toDispatch();

			String formulaCell = "";

			formulaCell = cellref.split("\\d")[0];
			System.out.println("formula cell ref=" + formulaCell);
			int index = Integer.parseInt(cellref.split("([A-Za-z]+)")[1]);
			System.out.println("index cell ref=" + index);
			index++;

			formulaCell += index;
			Dispatch.put(cell1, "Value", content);
			Dispatch.put(Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] { formulaCell }, new int[] { 1 })
					.toDispatch(), "Formula", "=TRIM(" + cellref + ")");

			Dispatch.call(wkbksingle, "SaveAs", targetFileName);
			Dispatch.call(component, "Quit");
			result = true;
		} catch (Exception e) {
			result = false;
		}

		return result;
	}
}