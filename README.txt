本jar使用方式:
1.编写类xxProjectExcelHandle.java 继承成ProjectExcelHandle 实现getParamsSql()方法
2.编写报表实现类XxxxHandleImpl.java 继承xxProjectExcelHandle.java
    实现接口IExcelHandle的 calculateParamsValueMap()方法
3.创建ReprotExcelManager对象,并执行manage(JSONObject webParamsJsonObj)方法,
  webParamsJsonObj 中必须含有 "exceltitle"--报表模板名称参数

备注: ConstsLocalTest.LOCALTEST=ture 表示开启本地测模式 正式jar需要关闭此模式