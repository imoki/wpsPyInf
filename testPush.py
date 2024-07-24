import requests 

# 唯一id用于python脚本定位表格位置，从而能获取传入参数
uniqueId = "github123456789"

 
# 推送 
def push(pushType, key): 
  if key != "" : 
      if pushType.lower() == "bark": 
        url = "https://api.day.app/" + key + "/github推送脚本运行正常" 
      elif pushType.lower()  == "pushplus": 
        url = "http://www.pushplus.plus/send?token=" + key + "&content=github推送脚本运行正常" 
      elif pushType.lower()  == "serverchan": 
        url = "https://sctapi.ftqq.com/" + key + ".send?title=运行结果&desp=github推送脚本运行正常" 
      else: 
        url = "https://api.day.app/" + key + "/github推送脚本运行正常" 
      response = requests.get(url) 
      print(response.text) 
 
 
if __name__ == "__main__": 
  print("这是一段github上拉取下来的推送测试代码") 
  key = xl("k3", sheet_name="CONFIG")[0][0] # 访问表格 
  print(key) 
  keyarry = key.split("&") 
  for i in range(len(keyarry)): 
    pushType = keyarry[i].split("=")[0] 
    key = keyarry[i].split("=")[1] 
    push(pushType, key) 