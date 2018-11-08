# Report_jxl 报表导出jar

#### 优点：
1.通过"模板+数据"的模式，快速构建报表  
2."模板"使得报表样式可以由具体的使用人员自定义，将程序员从繁琐的调整excel报表样式中解脱出来。  
3."数据"按tag映射数据，将数据和excel表格位置解耦，新增报表项时，不用依次调整单元格位置。  
4.其他你们补充下，多说点优点。：）  

#### 缺点：
- [ ] 1.读取模板IO较多（TODO 缓存模板）
- [ ] 2.按行，或者其他规则批量，动态地写入不智能，方法不够清晰（TODO 设计一套按行写入的方法）
- [ ] 3.没有图片插入，等其他元素插入的公用方法（TODO：这个用的较少，感兴趣童鞋可以补充进来。）
- [ ] 4.最大缺点是，应该有人做过这个jar了吧，只是我没找到，找到的童鞋希望发我个信息。

  
## 本jar使用方式:
1.编写类xxProjectExcelHandle.java 继承成ProjectExcelHandle 实现getParamsSql()方法  
2.编写报表实现类XxxxHandleImpl.java 继承xxProjectExcelHandle.java
    实现接口IExcelHandle的 calculateParamsValueMap()方法  
3.创建ReprotExcelManager对象,并执行manage(JSONObject webParamsJsonObj)方法,
  webParamsJsonObj 中必须含有 "exceltitle"--报表模板名称参数  
备注: ConstsLocalTest.LOCALTEST=ture 表示开启本地测模式 正式jar需要关闭此模式  
```
# 待补充
```
  
## TODO
最近在新的项目中用到了，做了些完善，有空再补充进来。

