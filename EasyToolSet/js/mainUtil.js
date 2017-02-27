/**
 * Created by hp on 2017/1/9.
 */
$(document).ready(function(){
    //new ActiveXObject("Excel.application");
    if(!Common.isIE()){
        //$('#upload').attr('disabled',true)
    }
    $("#submit").click(function(e){
        var path = $(e.target).prev().val()
        var reader = null
        try{
            reader = new ExcelReader();
            reader.openExcel(path);

            //
            var map =new Object()
            //读取全部数据
            var dataList = reader.readAll();
            //筛选数据
            for(var i =0;i<dataList.length;i++){
                if(map[dataList[i]]){
                    map[dataList[i]].push(dataList[i].content)
                }else{
                    map[dataList[i]] = [dataList[i].content]
                }
            }


        }catch (e){
            console.log(e)
            reader.close()
        }
    })

    //单元测试
    //$('#test').click(function(){
    //    var writer = new ExcelWriter("C:\\tesdsst.xlsx")
    //    writer.createExcel()
    //    writer.write(1,1,"ceshi").write(2,2,"hello").renameSheet('newname')
    //    writer.close()
    //})
    //
    //$('#test2').click(function(){
    //    var list = []
    //    list[10] = 100
    //    console.log(list)
    //})
})

/**
    普通方法
 */
var Common = {
    isIE:function(){
        if ((navigator.userAgent.indexOf('MSIE') >= 0)
            && (navigator.userAgent.indexOf('Opera') < 0)){
            return true
        }
        else{
            alert('请使用IE浏览器')
            return false
        }
    }
}
