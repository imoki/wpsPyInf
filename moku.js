/*
    作者: imoki
    仓库: https://github.com/imoki/wpsPyInf
    公众号：默库
    更新时间：20240720
    脚本：moku.js 粘贴到金山文档内时，请改名为“默库”。
    说明：注意！请将文档名和脚本名都起名为“默库”，脚本才能正常运行。
          1. 第一步，首次运行“默库”脚本（仓库中的“moku.js”）会生成wps表，请先填写好wps表的内容，只填wps_sid即可。
          2. 填写CONFIG表的内容。默认任务1用于消息推送测试，测试脚本是否正常，填写推送的key即可，如:bark=xxxx&pushplus=xxxx。(这一步可以跳过)
          3. 再运行一次“默库”脚本，此时你将收到推送通知，说明你操作正确，可正常使用了。(这一步可以跳过)
          4. 请在CONFIG表填写你自己写的python脚本和定时时间，然后运行一次“默库”脚本，即可按照配置好的来执行脚本，就不需要再管了。
*/

var sheetNameSubConfig = "wps"; // 分配置表名称
var sheetNameConfig = "CONFIG"
var sheetName = "默库"
var cookie = ""
var taskArray = []
var headers = ""
var count = "20" // 读取的文档页数
var excludeDocs = []
var onlyDocs = [] // 仅读取哪些文档
// 表中激活的区域的行数和列数
var row = 0;
var col = 0;
var maxRow = 100; // 规定最大行
var maxCol = 16; // 规定最大列
var workbook = [] // 存储已存在表数组
var colNum = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']

// 定时任务相关
var hourMin = 0
var hourMax = 23
var cron_type = "daily"
var day_of_month = 0
var day_of_week = 0
var task_id = 0 // 定时任务id
var asid = 0

// 注入脚本
var file_id = 0;
var script_id = 0;

function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d; );
}

// 激活工作表函数
function ActivateSheet(sheetName) {
    let flag = 0;
    try {
      // 激活工作表
      let sheet = Application.Sheets.Item(sheetName);
      sheet.Activate();
      // console.log("🥚 激活工作表：" + sheet.Name);
      flag = 1;
    } catch {
      flag = 0;
      console.log("🍳 无法激活工作表，工作表可能不存在");
    }
    return flag;
}

// 存储已存在的表
function storeWorkbook() {
  // 工作簿（Workbook）中所有工作表（Sheet）的集合,下面两种写法是一样的
  let sheets = Application.ActiveWorkbook.Sheets
  sheets = Application.Sheets

  // 清空数组
  workbook.length = 0

  // 打印所有工作表的名称
  for (let i = 1; i <= sheets.Count; i++) {
    workbook[i - 1] = (sheets.Item(i).Name)
    // console.log(workbook[i-1])
  }
}

// 判断表是否已存在
function workbookComp(name) {
  let flag = 0;
  let length = workbook.length
  for (let i = 0; i < length; i++) {
    if (workbook[i] == name) {
      flag = 1;
      console.log("✨ " + name + "表已存在")
      break
    }
  }
  return flag
}

// 创建表，若表已存在则不创建，直接写入数据
function createSheet(name) {
  // const defaultName = Application.Sheets.DefaultNewSheetName
  // 工作表对象
  if (!workbookComp(name)) {
    Application.Sheets.Add(
      null,
      Application.ActiveSheet.Name,
      1,
      Application.Enum.XlSheetType.xlWorksheet,
      name
    )
  }
}


// 获取wps_sid、cookie
function getWpsSid(){
  // flagConfig = ActivateSheet(sheetNameSubConfig); // 激活wps配置表
  // 主配置工作表存在
  if (1) {
    console.log("🍳 开始读取wps配置表");
    for (let i = 2; i <= 100; i++) {
      // 读取wps表格配置
      wps_sid = Application.Range("A" + i).Text; // 以第一个wps为准
      // name = Application.Range("H" + i).Text;
      
      excludeDocs = Application.Range("C" + i).Text.split("&")
      onlyDocs = Application.Range("D" + i).Text.split("&")

      break
    }
  }
  return wps_sid
  
  // filename = name
}



// 判断是否为xlsx文件
function juiceXLSX(name){
  let flag = 0
  let array= name.split(".") // 使用|作为分隔符
  if(array.length == 2 && (array[1] == "xlsx" || array[1] == "ksheet")){
    flag = 1
  }
  return flag 
}

// 判断是否为要排除文件
function juiceDocs(name){
  let flag = 0
  if((excludeDocs.length == 1 && excludeDocs[0] == "") || excludeDocs.length == 0){
    flag = 0
    // console.log("excludeDocs不符合")
  }else{
    for(let i= 0; i<excludeDocs.length; i++){
      if(name == excludeDocs[i]){
        flag = 1  // 找到要排除的文档了
        // console.log("找到要排除的文档了")
      }
    }
  }
  
  return flag 
}

// 判断是否为仅读取的文档
function juiceOnlyRead(name){
  let flag = 0  // 不读取
  if(onlyDocs == "@all"){
    flag = 1  // 所有都读取
    // console.log("所有都读取")
  }else{
    for(let i= 0; i<onlyDocs.length; i++){
      if(name == onlyDocs[i]){
        flag = 1  // 找到要读取的文档了
        // console.log("找到要读取的文档了")
      }
    }
  }
  
  return flag 
}

// 判断是否存在定时任务
function taskExist(file_id){
  url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/cron_tasks";
  // console.log(url)
  // 查看定时任务
  resp = HTTP.get(
    url,
    { headers: headers }
  );

  resp = resp.json()
  // console.log(resp)
  // list -> 数组 -> file_id、task_id、script_name，cron_detail->字典
  cronlist = resp["list"]
  sleep(3000)
  return cronlist
}


// 创建脚本
function createPyScript(url, headers){
  data = {"script_name": sheetName,"script":"","ext":"py"}
  let resp = HTTP.post(
    url,
    data,
    { headers: headers }
  );
  // {"id":""}
  resp = resp.json()
  id = resp["id"]

  return id
}

// 判断表格行列数，并记录目前已写入的表格行列数。目的是为了不覆盖原有数据，便于更新
function determineRowCol() {
  for (let i = 1; i < maxRow; i++) {
    let content = Application.Range("A" + i).Text
    if (content == "")  // 如果为空行，则提前结束读取
    {
      row = i - 1;  // 记录的是存在数据所在的行
      break;
    }
  }
  // 超过最大行了，认为row为0，从头开始
  let length = colNum.length
  for (let i = 1; i <= length; i++) {
    content = Application.Range(colNum[i - 1] + "1").Text
    if (content == "")  // 如果为空行，则提前结束读取
    {
      col = i - 1;  // 记录的是存在数据所在的行
      break;
    }
  }
  // 超过最大行了，认为col为0，从头开始

  // console.log("✨ 当前激活表已存在：" + row + "行，" + col + "列")
}

// 统一编辑表函数
function editConfigSheet(content) {
  determineRowCol();
  let lengthRow = content.length
  let lengthCol = content[0].length
  if (row == 0) { // 如果行数为0，认为是空表,开始写表头
    for (let i = 0; i < lengthCol; i++) {
      Application.Range(colNum[i] + 1).Value = content[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      Application.Range(colNum[j] + i).Value = content[i - 1][j]
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = content[i - 1][j]
    }
  }
}

// 创建wps表
function createWpsConfig(){
  createSheet(sheetNameSubConfig) // 若wsp表不存在创建wps表
  let flagExitContent = 1

  if(ActivateSheet(sheetNameSubConfig)) // 激活wps配置表
  {
    // wps表内容
    let content = [
      ['wps_sid', '任务配置表超链接', '文档id', 'Pyid', 'Asid', '定时任务id', ],
      ['此处填写wps_sid', "点击此处跳转到" + sheetNameConfig + "表" ,'', '', '', '' ]
    ]
    determineRowCol() // 读取函数
    if(row <= 1 || col < content[0].length){ // 说明是空表或只有表头未填写内容，或者表格有新增列内容则需要先填写
      // console.log(row)
      flagExitContent = 0 // 原先不存在内容，告诉用户先填内容
      editConfigSheet(content)
      // console.log(row)
      let name = "点击此处跳转到" + sheetNameConfig + "表"  // 'CRON'!A1
      let link = sheetNameConfig
      let link_name ='=HYPERLINK("#'+link+'!$A$1","'+name+'")' //设置超链接
      Application.Range("B2").Value = link_name
    }
  }

  return flagExitContent
  
}

// 创建CONFIG表
function createConfig(){
  createSheet(sheetNameConfig) // 若CONFIG表不存在则创建
  let flagExitContent = 1

  if(ActivateSheet(sheetNameConfig)) // 激活配置表
  {
    // CONFIG表内容
    // 推送昵称(推送位置标识)选项：若“是”则推送“账户名称”，若账户名称为空则推送“单元格Ax”，这两种统称为位置标识。若“否”，则不推送位置标识
    
    testPythonScript = "import requests\r\n\r\n# 推送\r\ndef push(pushType, key):\r\n  if key != \"\" :\r\n      if pushType.lower() == \"bark\":\r\n        url = \"https://api.day.app/\" + key + \"/运行正常\"\r\n      elif pushType.lower()  == \"pushplus\":\r\n        url = \"http://www.pushplus.plus/send?token=\" + key + \"&content=运行正常\"\r\n      elif pushType.lower()  == \"serverchan\":\r\n        url = \"https://sctapi.ftqq.com/\" + key + \".send?title=运行结果&desp=运行正常\"\r\n      else:\r\n        url = \"https://api.day.app/\" + key + \"/运行正常\"\r\n      response = requests.get(url)\r\n      print(response.text)\r\n\r\n\r\nif __name__ == \"__main__\":\r\n  print(\"这是一段推送测试代码\")\r\n  key = xl(\"k2\", sheet_name=\"CONFIG\")[0][0] # 访问表格\r\n  print(key)\r\n  keyarry = key.split(\"&\")\r\n  for i in range(len(keyarry)):\r\n    pushType = keyarry[i].split(\"=\")[0]\r\n    key = keyarry[i].split(\"=\")[1]\r\n    push(pushType, key)\r\n\r\n\r\n\r\n  \r\n"
    testKey = "bark=&pushplus=&ServerChan="
    let content = [
      ['任务的名称', '备注', '更新时间', '消息', '推送时间', '推送方式',  '是否通知', '是否加入消息池', '是否执行', '脚本', '脚本传入参数', '定时时间'],
      ['任务1', '随便填给自己看的', '', '' , '' , '@all' , '是', '否' , '是' , testPythonScript, testKey, '8:00' ],
      ['任务2', '任务3通知', '', '' , '' , '@all' , '是', '否' , '否' , '', '', '8:10' ],
      ['任务3', '任务3通知', '', '' , '' , '@all' , '是', '否' , '否' , '', '', '9:00' ],
    ]
    determineRowCol() // 读取函数
    if(row <= 1 || col < content[0].length){ // 说明是空表或只有表头未填写内容，或者表格有新增列内容则需要先填写
      // console.log(row)
      flagExitContent = 0 // 原先不存在内容，告诉用户先填内容
      editConfigSheet(content)
    }
  }

  return flagExitContent
  
}

// 获取file_id
function getFile(url){
  let flag = 0
  // 查看定时任务
  resp = HTTP.get(
    url,
    { headers: headers }
  );

  resp = resp.json()
  // console.log(resp)
  resplist = resp["list"]
  for(let i = 0; i<resplist.length; i++){
    roaming = resplist[i]["roaming"]
    // console.log(roaming)
    fileid = roaming["fileid"]
    name = roaming["name"]
    // 找到指定文档
    if(juiceXLSX(name) && sheetName == name.split(".")[0]){
      console.log("✨ 已找到" + sheetName + "文档")
      file_id = fileid
      flag = 1
      break;  // 找到就退出
    }
  }

  // console.log(taskArray)
  
  sleep(3000)
  return flag
}

// python脚本列表
function pyScriptList(file_id){
  let url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/script?ext=py"
  // console.log(url)
  // 查看定时任务
  let resp = HTTP.get(
    url,
    { headers: headers }
  );

  resp = resp.json()
  // console.log(resp)

  let list = resp["data"]
  sleep(3000)
  return list
}

// 判断是否存在某脚本，写入script_id
function existPythonScript(){
  let flagFind = 0
  list = pyScriptList(file_id)
  // console.log(list)
  // {
  //     "data": [
  //         {
  //             "id": "V7-xxxxx",
  //             "script_name": "a",
  //             "view_config": "",
  //             "update_at": ,
  //             "edit_permission": 1,
  //             "is_admin": true,
  //             "read_only": false,
  //             "creator_id": "",
  //             "creator_name": "",
  //             "create_time": ,
  //             "last_modifier_id": "",
  //             "last_modifier_name": "",
  //             "last_modify_time": 
  //         }
  //     ]
  // }
  if(list != undefined){
    if(list.length > 0){
      console.log("🎉 存在python任务")
      // console.log(list)
      for(let i = 0; i < list.length; i++){
        
        task = list[i]
        let id = task["id"]
        script_name = task["script_name"]

        // 查找是否有指定脚本
        if(script_name == sheetName){
          console.log("🎉 存在" + sheetName + "脚本")
          script_id = id
          // taskArray.push({
          //   "filename" : name,
          //   "fileid" : fileid,
          //   "script_id" : script_id,
          //   "script_name" : script_name,
          // })

          flagFind = 1
          break;
        }

    }
    
    }
  }

  return flagFind
}

// 执行脚本
function runScript(url, headers, script){
  let data = {"sheet_name":"task","script":script}
  // console.log(data)

  let resp = HTTP.post(
      url,
      data,
      { headers: headers },
  );
  resp = resp.json()
  // {"data":{"grant":{"need":[{"name":"http","open":true},{"name":"smtp","open":true}]}},"result":"ok"}
  // console.log(resp)
  let result = resp["result"]
  return result
}

// 修改定时任务
function putTask(url, headers, data, task_id, script_name){
  let flagResult = 0
  // console.log(url)
  // console.log(data)
  // console.log(headers)
  // console.log(task_id)
  if(task_id == "undefined" || task_id == null || task_id == ""){
    console.log("🎉 创建" + sheetName + "定时任务")
    // 创建定时任务
    resp = HTTP.post(
      url,
      data,
      { headers: headers }
    );
  }else{
    // 修改定时任务
    resp = HTTP.put(
      url + "/" + task_id,
      data,
      { headers: headers }
    );
  }

  resp = resp.json()
  // console.log(resp)

  // {"result":"ok"}
  // {"errno":10000,"msg":"not find script","reason":"","result":"InvalidArgument"}
  // {"errno":10000,"msg":"value of hour is out of 24's bounds","reason":"","result":"InvalidArgument"}
  // {"errno":10000,"msg":"status not allow","reason":"","result":"InvalidArgument"}
  // {"task_id":""}
  result = resp["result"]
  if(result == "ok"){
    console.log("🎉 " + script_name + " 任务时间调整成功")
    flagResult = 1
  }else{
    task_id = resp["task_id"]
    // console.log(task_id)
    if(task_id != "undefined" || task_id == null){
      // console.log("task_id不为空")
      return task_id
    }else{
      msg = resp["msg"]
      console.log("📢 " , msg)
    }

  }
  sleep(5000)
  return flagResult
}

// 返回当前月和星期几
function getMonthWeek(){
  let mw = []
  let date = new Date();
  let weekdayIndex = date.getDay(); // getDay()返回的是0（星期日）到6（星期六）之间的一个整数
  mw[0] = date.getDate().toString()
  mw[1] = weekdayIndex.toString()
  // if(mw[1] == "0"){ // 星期日返回7
  //   mw[1] = 7
  // }
  return mw
}

// 获取系统时分
function getsysHM(){
  let syshm = []
  let currentDate = new Date();
  // 获取小时（注意：getHours() 返回的是 0-23 之间的数）
  let syshours = currentDate.getHours();
  // 获取分钟（getMinutes() 返回的是 0-59 之间的数）
  let sysminutes = currentDate.getMinutes();
  // // 获取秒（getSeconds() 返回的是 0-59 之间的数）
  // let sysseconds = currentDate.getSeconds();
  // 如果需要两位数的小时、分钟或秒，可以使用 padStart() 方法来填充前导零
  syshours = syshours.toString().padStart(2, '0');
  sysminutes = sysminutes.toString().padStart(2, '0');
  // sysseconds = seconds.toString().padStart(2, '0');

  syshm[0] = parseInt(syshours)
  syshm[1] = parseInt(sysminutes)
  return syshm
}

// 获取airScipt_id
function getAsId(){
  url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/script"
  // console.log(url)
  // 创建定时任务
  resp = HTTP.get(
    url,
    { headers: headers }
  );

  resp = resp.json()
  // console.log(resp)
  let list = resp["data"]
  for(let i = 0; i<list.length; i++){
    let name = list[i].script_name
    if(name == sheetName){
      asid =  list[i].id
      console.log("✨ 写入找到Asid:" + asid)
    }
  }

  sleep(5000)
}

// 数组字符串转整形
function arraystrToint(array){
  let result = []
  for(let i=0; i<array.length; i++){
    result.push(parseInt(array[i]))
  }
  return result
}

// 数组升序排序
function arraySortUp(value){
  value.sort(function(a, b) {
    return a - b; // 升序排序
  });
  return value
}

// 数组-字典字符串转整形
function dictarraystrToint(array){
  let result = []
  for(let i=0; i<array.length; i++){
    result.push({
        "hour" : parseInt(array[i]["hour"]),
        "minute" : parseInt(array[i]["minute"]),
        "pos" : array[i]["pos"],
        // "flagExec" : array[i]["flagExec"],
      })
  }
  return result
}

// 数组-字典升序排序，按分
function dictarraySortUpMinute(value){
  value.sort(function(a, b) {
    // console.log(a, b)
    return a["minute"] - b["minute"]; // 升序排序
  });
  return value
}

// 数组-字典升序排序，按时
function dictarraySortUpHour(value){
  value.sort(function(a, b) {
    // console.log(a, b)
    return a["hour"] - b["hour"]; // 升序排序
  });
  return value
}

// 运行任务
function runtask(){
  // 根据task表运行任务

  // 判断是否有CONFIG表
  flagConfig = ActivateSheet(sheetNameConfig); // 激活cron配置表
  // 主配置工作表存在
  if (flagConfig == 1) {
    
    
    // 执行逻辑：先设置新定时， 再执行py脚本
    
    // 找到下一次执行脚本的定时
    // 获取系统时间，比对时间，找到最接近的靠后（右边）的时间，相对则最优先


    // 处理任务列表的时间，记录时分及位置
    // 排序时间
    // 找到靠右最接近的时间，获得此位置。设置为定时。
    
    // 读取生成任务列表
    let pos = 0
    let hourarry = []
    for(let t = 0; t < maxRow; t++){
      pos = t + 2
      script_name = Application.Range(colNum[0] + pos).Text 
      if(script_name == ""){
        break
      }
      flagExec = Application.Range(colNum[8] + pos).Text 
      hm = Application.Range(colNum[11] + pos).Text   // 时间 例如：8:10
      hour = hm.split(":")[0],
      minute = hm.split(":")[1]
      if(flagExec == "是")  // 是否执行，是的才加入任务
      {
        dict = {
          "hour" : hour,
          "minute" : minute,
          "pos" : pos,
          // "flagExec" : flagExec,
        }
        taskArray.push(dict)

      }

    }
    // console.log(taskArray)
    taskArray = dictarraySortUpMinute(taskArray) // 升序排序，按分
    taskArray = dictarraySortUpHour(taskArray)  // 升序排序， 再按时
    taskArray = dictarraystrToint(taskArray)  // 转整形
    // console.log(taskArray)

    // 用于定时的时分
    hour = 0
    minute = 0
    let flagChange = 0
    pos = 0 // 记录位置
    let index = 0 // 计入任务索引，下标

    let syshm = getsysHM()  // 获取系统时间，时分
    let sysminuteSum = syshm[0] * 60 + syshm[1]
    // console.log(sysminuteSum)
    // 查找靠右第一个
    for(let j=0; j < taskArray.length; j++){
      let hourExpect = taskArray[j]["hour"]
      let minuteExpect = taskArray[j]["minute"]
      
      // 用总分钟比较生成值：hour*60 + minute = minuteSum
      minuteSum = hourExpect*60 + minuteExpect
      // console.log("任务时分：", hourExpect, ":" , minuteExpect)
      // console.log("任务总分钟", minuteSum)
      if(sysminuteSum < minuteSum){
        pos = taskArray[j]["pos"]
        index = j
        // 取第一个遇到比原先大的值，就取它
        hour = hourExpect
        // console.log(String(minuteExpect))
        if(String(minuteExpect) == "NaN"){
          // console.log("minuteExpect 为空")
        }else{
          minute = minuteExpect
        }
        
        flagChange = 1
        break
      }
    }
    // console.log(taskArray)
    // 查找最小值
    if(!flagChange){  // 如果时间没变动， 说明当前时间已经时最大了，则置为最小值
      // console.log("时间没变动， 置为最小值")
      pos = taskArray[0]["pos"]
      index = 0
      let hourExpect = taskArray[0]["hour"]
      let minuteExpect = taskArray[0]["minute"]
      hour = hourExpect
      // console.log(String(minuteExpect))
      if(String(minuteExpect) == "NaN"){
        // console.log("minuteExpect 为空")
      }else{
        minute = minuteExpect
      }

    }

    // console.log("任务索引：" , index)
    // console.log("位置：" , pos)
    // console.log("时分：", hour, ":" , minute)
    
    // 定时任务修改
    // 进行时间修改，不存在则修改
    let nw = getMonthWeek()
    day_of_month = nw[0]
    day_of_week = nw[1]
    
    // console.log(url)
    url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/cron_tasks";
    // 写入新的定时任务
    // console.log(task_id)
    if(task_id == "undefined" || task_id == null || task_id == "" || task_id == 0 || task_id == undefined){
      // 无定时任务，直接写新定时任务
      data = {
        "id": file_id.toString(),
        "script_id": asid,
        "cron_detail": {
            "task_type": "cron_task",
            "cron_desc": {
              "cron_type": cron_type,
              "day_of_month": day_of_month.toString(),
              "day_of_week": day_of_week.toString(),
              "hour" : hour.toString(),
              "minute": minute.toString()
            }
        }
      }
    }else{
      // 已有定时任务
      
      data = {
        "id": file_id.toString(),
        "script_id": asid,
        "cron_detail": {
            "task_type": "cron_task",
            "cron_desc": {
                "cron_type": cron_type,
                "day_of_month": day_of_month.toString(),
                "day_of_week": day_of_week.toString(),
                "hour" : hour.toString(),
                "minute": minute.toString()
            }
        },
        "task_id": task_id,
        "status": "enable"
      }
    }
    // console.log(url)
    // console.log(data)

    console.log("✨ 现定时任务：" , script_name, " 定时时间：", hour,":",  minute)
    let flagResult = putTask(url, headers, data, task_id, script_name)
    if(flagResult != 1){ // 返回的是task_id，记录下task_id
      // 写入task_id，定时任务id
      ActivateSheet(sheetNameSubConfig) // 激活wps配置表
      // console.log(flagResult)

      task_id = flagResult
      // console.log(task_id)
      console.log("✨ 写入定时任务id")
      Application.Range("F2").Value = task_id
      ActivateSheet(sheetNameConfig) // 激活CONFIG配置表
    }

    console.log("✨ 已将下一个任务安排进定时任务中")

    // 运行定时任务
    // 取设定时的前一个任务来运行，即当前应该运行的任务
    if(index <= 0){
      pos = taskArray[0]["pos"]
    }else{
      pos = taskArray[index - 1]["pos"]
    }
    // console.log("✨ 执行当前任务位置：" , pos)

    // 安排下一个任务进定时任务中

    // 调用执行py脚本
    console.log("✨ 已获取到" + sheetNameConfig + "表，开始注入任务")
    // let pos = 2
    script = Application.Range(colNum[9] + pos).Text 
    // console.log(script)
    script_name = Application.Range(colNum[1] + pos).Text 
    // console.log(script_name)
    // 执行脚本
    // file_id = parseInt(file_id)
    url = "https://www.kdocs.cn/api/aigc/pyairscript/v2/" + file_id + "/script/" + script_id + "/exec"
    // console.log(url)
    let result = runScript(url, headers, script)
    
    if(result == "ok"){
      console.log("✨ " + script_name + " 已执行")
    }else{
      console.log("📢 " + script_name + "执行失败")
    }

    sleep(3000)
    
  }else{
    createSheet(sheetNameConfig)  
    console.log("📢 请先填写" + sheetNameConfig + "表中的内容")
  }


}

// 权限允许
function permissionOn(){
  url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/script/" + script_id + "/permission"
  resp = HTTP.post(
    url,
    data,
    { headers: headers }
  );
  // 404 page not found
  // console.log(resp.text())
  resp = resp.json()
  // console.log(resp)

  result = resp["result"]
  if(result == "ok"){
    console.log("🎉 已允许网络请求")
  }else{
     console.log("📢 请手动赋予网络API权限，并点击运行，再点击允许网络请求")
  }
  sleep(5000)
}

// 赋予网络api权限
function change_permission_config(){
  url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/script/" + script_id
 
  data = {
      "change_permission_config": true,
      "id": script_id,
      "permission_config": {
          "ks_drive": {
              "open": false,
              "allow_open_all_file": false,
              "allow_open_files": null
          },
          "http": {
              "open": true,
              "allow_all_host": true,
              "allow_hosts": null
          },
          "smtp": {
              "open": true,
              "allow_all_email": true,
              "allow_emails": null
          },
          "sql": {
              "open": false
          }
      }
  }

  //  console.log(url)
  //  console.log(data)
  resp = HTTP.put(
    url,
    data,
    { headers: headers }
  );
  // 404 page not found
  // console.log(resp.text())
  resp = resp.json()
  // console.log(resp)

  result = resp["result"]
  if(result == "ok"){
    console.log("🎉 成功赋予网络API权限")
  }else{
     console.log("📢 请手动赋予网络API权限")
  }
  sleep(5000)
  // return flagResult
}

// 初始化，无文档id和脚本id的时候使用
function init(){
  // try{
  //   Application.Sheets.Item(sheetName).Delete()  // 为了获得最新数据，删除表
  //   storeWorkbook()
  // }catch{
  //   console.log("🍳 不存在" + sheetName + "表，开始进行创建")
  // }
  // 判断是否以前已写入数据
  if(ActivateSheet(sheetNameSubConfig)) // 激活wps配置表
  {
    // 定时任务id
    task_id = Application.Range("F2").Value
    
    // 读取文档id
    file_id = Application.Range("C2").Value
    // console.log(file_id)
    if(file_id != "" && file_id != 0 && file_id != null){
      console.log("✨ 已读取文档id")
    }else{
      // 无文档id，则写入文档id

      // 获取文档id
      url = "https://drive.kdocs.cn/api/v5/roaming?count=" + count  // 只对前20条进行判断
      let flagFindFileid = getFile(url)
      if(flagFindFileid == 0){
        console.log("📢 请将本文档名称更改为 " + sheetName + " 然后再运行一次脚本")
      }else{
        // 有文档id了
        // 写入文档id
        console.log("✨ 写入文档id")
        let pos = 2
        Application.Range(colNum[2] + pos).Value = file_id
      }
    }
    
    // console.log(file_id)
    if(file_id != "" && file_id != 0){
      // 读取脚本id
      let i = 2
      script_id = Application.Range("D" + i).Text
      if(script_id != "" && script_id != 0){
        console.log("✨ 已获取到" + sheetName + "脚本")
      }else{
        // 无指定脚本，可能是第一次运行或清空了id，则进行数据写入以及py脚本创建

        // 若是清空了id，脚本还存在，则不创建脚本仅写入id
        let flagFind = existPythonScript()  // 判断是否存在指定脚本
        if(flagFind){
          // 说明已有所需py脚本
          Application.Range(colNum[3] + "2").Value = script_id
          console.log("✨ 已有" + sheetName + "脚本，写入最新id")
        }else{
          // 无指定的脚本，是第一次运行，则进行数据写入以及py脚本创建
          
          // 第一次运行  
          url = "https://www.kdocs.cn/api/v3/ide/file/" +file_id + "/script"
          script_id = createPyScript(url, headers)  // 创建脚本
          // console.log(script_id)

          // 写入脚本id
          let pos = 2
          Application.Range(colNum[3] + pos).Value = script_id
          console.log("✨ 已创建" + sheetName + "脚本")
          console.log("✨ 请将" + sheetName + "脚本加入定时任务")
        }

        // 赋予网络api权限
        change_permission_config()

        // 允许网路请求
        permissionOn()

      }

      asid = Application.Range("E2").Value
      // console.log(asid)
      // 如果没有asid
      if(asid == "" || asid == "undefined" || asid == null){
        getAsId()
        Application.Range("E2").Value = asid
      }
      // console.log(asid)
    }

  }
    

  
  // // 获取file_id
  // url = "https://drive.kdocs.cn/api/v5/roaming?count=" + count  // 只对前20条进行判断
  // let flagFind = getFile(url)
  // if(flagFind){
  //   // 说明已创建所需py脚本
  //   console.log("✨ 已有" + sheetName + "脚本")
  // }else{
  //   // 无指定脚本，可能是第一次运行，则进行数据写入以及py脚本创建

  //   // 创建脚本
  //   url = "https://www.kdocs.cn/api/v3/ide/file/xxx/script"
  //   let id = createPyScript(url, headers)
  //   console.log(id)

  //   writeTask()
  //   console.log("✨ 已完成对" + sheetName + "表的写入，请到" + sheetName + "表进行配置")
  //   console.log("✨ 然后将" + sheetName + "脚本加入定时任务，即可自动调整定时任务时间")
  // }

}

function main(){
  storeWorkbook()
  let flagExitContent = createWpsConfig()
  if(flagExitContent == 0){
    console.log("📢 请先填写wps表，然后再运行一次此脚本")
    createConfig()  // 第一次运行时，创建CONFIG表
  }else{
    wps_sid = getWpsSid() // 获取wps_sid
    cookie = "wps_sid=" + wps_sid // 获取cookie
    // console.log(excludeDocs)

    headers = {
      "Cookie": cookie,
      "Content-Type" : "application/json",
      "Origin":"https://www.kdocs.cn",
      "Priority":"u=1, i",
    }
    
    
    // 获取定时任务,生成CRON定时任务表
    init()

    // 执行脚本
    runtask()
  }

}

main()