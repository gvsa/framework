package framework;

import org.globallogic.auto.fwk.windows.WindowsUtility;
import org.junit.Assert;
import org.junit.Test;

import com.jacob.com.ComThread;

public class WindowsApplicationTest {

	private boolean result;

	@Test
	public void scenario() {
		// this needs to be done to initialize the COM handles to current ones
		ComThread.InitSTA();

		result = WindowsUtility.writeToExcel("gvs1.xlsx", "new updated content=    this is some new test content    ");

		Assert.assertTrue("Writing to Excel failed!", result);

		ComThread.quitMainSTA();

	}

}
