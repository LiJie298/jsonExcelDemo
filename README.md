# jsonExcelDemo
使用技术（jar）
--
        * poi
        * springboot
        * fastJson
主要功能
---
        1、编写配置文件，导出excle文件。支持单sheet文件导出，多sheet 文件导出；导出文件根据json数据自动格式化；
使用说明
--
        1、配置文件参考resource目录下的 excelFile.properties文件。
            其中
            demo1 (单sheet导出)
            title对应生成Excel的文件表头,data:是导入文件对应的字段（这版没有加上传模块，这个值目前没有什么作用）,imdata:导出文件的json对应字段;
       demo2（多个sheet导出）
使用过程
--
![postman](https://github.com/LiJie298/jsonExcelDemo/raw/master/dis/postman.png)
![result](https://github.com/LiJie298/jsonExcelDemo/raw/master/dis/excel.png)

测试数据
--
        1、json 转换成 excel
        参数：{
             "data": [
                 {
                     "ps": "123123123",
                     "orderId": 123123123,
                     "contractNum": 123123123,
                     "taskId": 123123123,
                     "customerName": "123123123",
                     "expOrderName": "123123123"
                 },
                 {
                     "ps": "123123123",
                     "orderId": 123123123,
                     "contractNum": 123123123,
                     "taskId": 123123123,
                     "customerName": "123123123",
                     "expOrderName": "123123123"
                 },
                 {
                     "ps": "123123123",
                     "orderId": 123123123,
                     "contractNum": 123123123,
                     "taskId": 123123123,
                     "customerName": "123123123",
                     "expOrderName": "123123123"
                 },
                 {
                     "ps": "123123123",
                     "orderId": 123123123,
                     "contractNum": 123123123,
                     "taskId": 123123123,
                     "customerName": "123123123",
                     "expOrderName": "123123123"
                 },
                 {
                     "ps": "123123123",
                     "orderId": 123123123,
                     "contractNum": 123123123,
                     "taskId": 123123123,
                     "customerName": "123123123",
                     "expOrderName": "123123123"
                 },
                 {
                     "ps": "123123123",
                     "orderId": 123123123,
                     "contractNum": 123123123,
                     "taskId": 123123123,
                     "customerName": "123123123",
                     "expOrderName": "123123123"
                 },
                 {
                     "ps": "123123123",
                     "orderId": 123123123,
                     "contractNum": 123123123,
                     "taskId": 123123123,
                     "customerName": "123123123",
                     "expOrderName": "123123123"
                 }
             ],
             "msg": "success",
             "status": 1
         }
        2、excel转换成json
        参数：
             "file":"",
             "fileType":"demo1",
             "rowNum":"2"
         
