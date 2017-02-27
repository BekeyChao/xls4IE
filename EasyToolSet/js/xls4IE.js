/**
 * xls4IE.js version 0.0.5
 * @author bekey
 * Update: 2017/1/11.
 * 仅供IE浏览器使用,不支持Edge
 * 需要用户配置相应的安全设置,安装office软件,并支持Active控件
 * 纯本地操作excel,适合完成本地自动化工具
 */

/**
 * Excel数据单元格
 * @params props Object 用于初始化
 */
function XlsData(props){
    this.row = props.row;   //数据所在行
    this.column = props.column; //数据所在列
    this.content = props.content;   //数据内容  目前仅支持文本数据
    //this.sheet = props.sheet;   //数据所在工作表名称  (因为index容易改变)
    //其他属性可以看情况增加

    /*
     * 返回是否匹配条件,全匹配返回true
     * selector 匹配条件{row:matcher,column:matcher,content:matcher} 仅传一个条件视为内容匹配
     * row,column ->   >num 是否大于  <num 是否小于 ==num 是否等于....  如果为多条件 用&&连接不超过1个
     * content -> 请传入正则表达式
     */
    this.isMatch = function(selector){
        if(selector.row){
            if(!this.rowIsMatch(selector.row)) return false
        }
        if(selector.column){
            if(!this.columnIsMatch(selector.column)) return false
        }
        if(selector.content){
            return this.content.test(selector.content)
        }else{
            if(selector.row || selector.column) return true
            return this.content.test(selector)
        }
    }
    this.rowIsMatch = function(selector){
        if(selector.indexOf("&&") > -1){
            return this.rowIsMatch(selector.split("&&")[0]) && this.rowIsMatch(selector.split("&&")[1])
        }else{
            var row = this.row
            return eval(row+selector)
        }
    }
    this.columnIsMatch = function(selector){
        if(selector.indexOf("&&") > -1){
            return this.columnIsMatch(selector.split("&&")[0]) && this.columnIsMatch(selector.split("&&")[1])
        }else{
            return eval(this.column+selector)
        }
    }
}

/**
 * 数据集合类,用以筛选数据等操作
 * 每一行 为一个数据对象
 * //TODO 问题挺多
 */
function XlsDataSet(){
    //以二维数组结构存储
    this.dataSet =null;
    //添加数据
    this.addData = function(xlsData){
        if(!this.dataSet["r"+xlsData.row])
            this.dataSet["r"+xlsData.row]=[]
        this.dataSet["r"+xlsData.row].push(xlsData)
    }
    //获得数据
    this.getOne = function(row,col){
        if(!this.dataSet['r'+row]) return null
        if(!this.dataSet['r'+row].column) return null
        return
    }
    //获取指定列所有非空数据
    this.getRowContent = function(row){
        if(!this.dataSet['r'+row]) return null
    }
}

/**
 * excel读取控件,为防止误操作不设改写功能
 * @parent ExcelBase
 */
function ExcelReader(){
    //继承ExcelBase类
    this.parent = ExcelBase;
    this.parent()
    delete this.parent
    //打开指定路径excel,并返回选定活动工作表
    this.openExcel = function(filePath){
        this.filePath =filePath
        var oXL = ExcelCommon.newActiveX()
        if(!oXL) return false
        this.oXL = oXL
        try{
            this.oWB = oXL.Workbooks.Open(this.filePath);
            this.changeSheet(1)
        }catch (e){
            throw e.message
        }
    }
    //读取当前活动工作表的全部内容,并返回非空值xlsData列表,问题挺多暂不能用
    this.readAllUable = function(){
        var dataSet = new XlsDataSet()
        var onValue;
        for(var i =1;i<=this.oSheet.UsedRange.Rows.Count;i++){
            //dataSet[i] = []
            for(var j =1;j<=this.oSheet.UsedRange.Columns.Count;j++ ){
                onValue = this.oSheet.Cells(i,j).Value;
                if(onValue){
                    dataSet.addData(new XlsData({row:i,column:j,content:onValue}))
                }
            }
        }
        return dataSet
    }
    this.readAll = function(){
        var list = []
        var content;
        for(var i =1;i<=this.oSheet.UsedRange.Rows.Count;i++){
            for(var j =1;j<=this.oSheet.UsedRange.Columns.Count;j++ ){
                content = this.oSheet.Cells(i,j).Value;
                if(content){
                    list.push(new XlsData({row:i,column:j,content:content}))
                }
            }
        }
        return list
    }
}

/**
 * excel写入控件
 * @parent ExcelBase
 */
function ExcelWriter(){
    //继承ExcelBase类
    this.parent = ExcelBase;
    this.parent()
    delete this.parent
    //新建Excel表格
    this.createExcel = function(filePath){
        this.filePath=filePath
        this.oXL = ExcelCommon.newActiveX()
        this.oWB = this.oXL.Workbooks.Add; //新增工作簿
        this.oWB.Worksheets(1).select();   //创建工作表
        this.oSheet = this.oWB.ActiveSheet;
        this.oSheet.SaveAs(this.filePath);
        //this.oXL.Visible = true;    //设置Excel可见
    }
    //指定行列写入内容,返回列表操作本身
    this.write = function(row,col,content){
        this.oSheet.Cells(row,col).Value = content;
        return this
    }
    //修改当前活动工作表,如果没有则创建
    this.changeSheet = function(sheetNo,sheetName){
        if(sheetNo <= this.oWB.Sheets.Count){
            this.oWB.worksheets(sheetNo).select();
            this.oSheet = this.oWB.ActiveSheet;
            return this
        }else{
            return this.addSheet(sheetName)
        }
    }
    //在第一个位置添加工作表,并选定,sheetName新建表名
    this.addSheet = function(sheetName){
        this.oWB.Sheets.Add()
        if(sheetName)
            return this.changeSheet(1).renameSheet(sheetName)
        return this.changeSheet(1)
    }
    //重命名当前活动工作表
    this.renameSheet = function(name){
        this.oSheet.Name = name
        return this
    }
    //写入一行数据
    this.writeLine = function(dataSet){

    }
    //写入一组数据 xlsData格式数组 --> [{row:row,column:col,content:content},{...}]
    this.writeArray = function(xlsDataList){
        for(var i =0;i<xlsDataList.length;i++){
            this.write(xlsDataList[i].row,xlsDataList[i].column,xlsDataList[i].content)
        }
    }
}

/**
 * 父类,定义基础方法与属性,请勿直接构建父类
 * @constructor
 */
function ExcelBase(){
    this.filePath = null //文件打开的路径
    this.oWB = null //工作簿
    this.oXL = null //excel进程
    this.oSheet = null //活动的工作表
    //关闭连接,操作结束后务必执行close或userControl方法之一
    this.close = function(){
        ExcelCommon.close(this.oWB,this.oXL)
    }
    //设置可见,控制权交由用户,操作结束后务必执行close或userControl方法之一
    this.userControl = function(){
        ExcelCommon.userControl(this.oWB,this.oXL)
    }
    //修改当前活动工作表,sheet工作表序列
    this.changeSheet = function(sheetNo){
        if(sheetNo <= this.oWB.Sheets.Count){
            this.oWB.worksheets(sheetNo).select();
            this.oSheet = this.oWB.ActiveSheet;
            return this
        }else{
            throw "大于表存在数量,修改失败"
        }
    }
    //返回按序列所有的Sheet名,第0位为null,后面按顺序排列
    this.findAllSheet = function(){
        var list = [null]
        for(var i=1,length = this.oWB.Sheets.Count;i<=length;i++){
            var name = this.oWB.Sheets(i).Name
            list.push(name)
        }
        return list
    }
    //获取当前操作信息,方便调试
    this.getStatus = function(){
        var status = this.filePath+","
        if(this.oWB)
            status += this.oWB.Name+","
        if(this.oSheet)
            status += this.oSheet.Name
        return status
    }

}

/**
 * excel读写控件 //TODO 不能使用
 *
 */
function ExcelServer(){

    //创建工作空间
    this.createWorkSpace = function(filePath){

    }
}

/**
 * 装饰器  用以调节页面格式\字体等内容
 * //TODO
 */
function XlsDecorator(){

}

/**
 * 自动化工具 用以自动打印等工作
 * //TODO
 */
function XlsAutoTool(){

}

/**
 * excel控件控制方法
 */
var ExcelCommon ={
    //创建一个新的activeX对象
    newActiveX:function(){
        var oXL = null;
        try{
            oXL = new ActiveXObject("Excel.application");
        }catch (e){
            throw e.message
        }
        if(!oXL){
            throw "创建Excel文件失败，可能是您的计算机上没有正确安装Microsoft Office Excel软件或浏览器的安全级别设置过高！"
        }
        return oXL
    },

    //关闭连接方法,避免占用资源
    close:function(oWB,oXL){
        if(oWB){
            oWB.Save()
            oWB.close()
            oWB=null
        }
        if(oXL){
            //oXL.UserControl = true  //excel交由用户控制
            oXL.Quit()  //Excel进程结束
            oXL = null
        }
    },
    //excel设置可见,控制器交由用户
    userControl: function(oWB,oXL){
        if(oWB){
            oWB.Save()
            oWB = null
        }
        if(oXL){
            oXL.UserControl = true  //excel交由用户控制
            oXL.Visible = true;
            oXL = null
        }
    }
}
