package chengjs.report_jxl;

import org.junit.Assert;
import org.junit.Test;

import chengjs.report_jxl.handle.IExcelHandle;

public class ReportExcelFactoryTest {

  @Test
  public void testCreate() throws Exception {
    String demoExcel = "demoExcel";
    IExcelHandle handle = ReportExcelFactory.create(demoExcel);
    Assert.assertNotNull(handle);
  }

}










