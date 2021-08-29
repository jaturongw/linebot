var ss = SpreadsheetApp.openById("1dCj2ePRSX3MruA0bQP9jBZcAh3THoLQB6g_R7HqGlPQ");
var indexsheet = ss.getSheetByName("indexto");

function doPost(e){

    var dialogflowDATA = JSON.parse(e.postData.contents);
    var message = dialogflowDATA.originalDetectIntentRequest.payload.data.message.text;
    var indexArray = indexsheet.getRange(2, 1, indexsheet.getLastRow(), indexsheet.getLastColumn()).getValues();
    var replymsg = '';
    var sheetName = "systemstatus";
    var sheet = ss.getSheetByName(sheetName);
    var tableArray = sheet.getRange(2, 1, sheet.getLastRow(),sheet.getLastColumn()).getValues();

    for (var i = 0; i < indexArray.length; i++) {

        if (indexArray[i][0] == message) {
            i = i + 2;

            var sheetName = indexsheet.getRange(i,2).getValue();
            var sheet = ss.getSheetByName(sheetName);
            var tableArray = sheet.getRange(2, 1, sheet.getLastRow(),sheet.getLastColumn()).getValues();

        }
    }

            for (var i = 0; i < tableArray.length; i++) {

            if (tableArray[i][0] == message) {
            i = i + 2;

            var replymsg = replayform(sheet,sheetName,i);
                    

            var replyJSON = 
            ContentService.createTextOutput(JSON.stringify(replymsg)).setMimeType(ContentService.MimeType.JSON);
            return replyJSON;
        }
    }
}

function showsystest(sheetin,row)
  {
      var replyccm = {
                 "fulfillmentMessages": [
                  {
                    "platform": "line",
                    "type": 4,
                    "payload": {
                        "line": {
                            "type": "text",
                            "text": sheetin.getRange(row, 3).getValue()
                        }
                    }
                }]
            } 
    return replyccm;
  }

function showsystemstatus(sheetin,row)
{
      var mainheader = "";
      var h1 = "สถานะบริการ 3BB GIGATV";
      var headend1 = "ทดสอบ";
      var stb1= "ทดสอบ";
      var system1 = "ทดสอบ";
      var network1 = "ทดสอบ";
      var internet1 = "ทดสอบ";
      var h2 = "สถานะการให้บริการ Aplication ต่างๆ";
      var gigatvapp = "ทดสอบ";
      var gigatvweb = "ทดสอบ";
      var hbogo = "ทดสอบ";
      var h3 = "ภาพรวมการให้บริการ";
      var reportdatein = "ทดสอบ";
      var report = "ทดสอบ";;
      var reportdate = "ทดสอบ";

      //var mainheader = "";
      // h1 = "สถานะบริการ 3BB GIGATV";
      //var headend1 = sheetin.getRange(row, 5).getValue();
      //var stb1= sheetin.getRange(row, 6).getValue();
      //var system1 = sheetin.getRange(row, 8).getValue();
      //var network1 = sheetin.getRange(row, 8).getValue();
      //var internet1 = sheetin.getRange(row, 9).getValue();
      //var h2 = "สถานะการให้บริการ Aplication ต่างๆ";
      //var gigatvapp = sheetin.getRange(row, 10).getValue();
      //var gigatvweb = sheetin.getRange(row, 11).getValue();
      //var hbogo = sheetin.getRange(row, 12).getValue();
      //var h3 = "ภาพรวมการให้บริการ";
      //var reportdatein = sheetin.getRange(row, 2).getValue();
      //var report = sheetin.getRange(row, 2).getValue();
      // reportdate = Utilities.formatDate(new Date(reportdatein),"GMT+7","dd/MM/yyy");
      

      var replysys = {
                 "fulfillmentMessages": [
                  {
                    "platform": "line",
                    "type": 4,
                    "payload": {
                        "line":  {  
                          "type": "flex",
                          "altText": "this is a flex message",
                          "contents": 
                        {
                          "type": "bubble",
                          "size": "giga",
                          "body": {
                            "type": "box",
                            "layout": "vertical",
                            "backgroundColor": "#B4EEF3FF",
                            "contents": [
                              {
                                "type": "text",
                                "text": ":: สรุปสถานะการให้บริการ :: "+ reportdate,
                                "weight": "bold",
                                "size": "md",
                                "align": "start",
                                "color": "#0000ff"
                              },
                              {
                                "type": "text",
                                "text":  h1,
                                "weight": "bold",
                                "size": "sm",
                                "align": "center",
                                "color": "#ff0000"
                              },
                              {
                                "type": "text",
                                 "size": "sm",
                                "text": "ช่องรายการ LIVE TV : "+ headend1
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "3BB GIGATV BOX :  "+ stb1
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "SYSTEM & SERVER : "+ system1
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "NETWORK "+ network1
                              },
                               {
                                "type": "text",
                                "size": "sm",
                                "text": "3BB Internet "+ internet1
                              },
                              {
                                "type": "text",
                                "text":  h2,
                                "weight": "bold",
                                "size": "sm",
                                "align": "center",
                                "color": "#ff0000"
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "3BB GIGATV App : "+ gigatvapp
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "3BB GIGATV Web : "+ gigatvweb
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "HBO GO : "+ hbogo
                              },
                              {
                                "type": "text",
                                "text":  h3,
                                "weight": "bold",
                                "size": "sm",
                                "align": "center",
                                "color": "#ff0000"
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": report
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "วันที่รายงาน : "+ reportdate
                              }
                                      ]
                                    }
                                  }
                                }
                              }
                            }]
                            }
 
    return replysys;
  }


function showccm(sheetin,row)
  {
      //var mainheader = "ทดสอบ";
      //var h1 = "ข้อมูลลูกค้า";
      //var account1 = "ทดสอบ";
      //var customer1 = "ทดสอบ";
      //var package1 = "ทดสอบ";
      //var loginid1 = "ทดสอบ";
      //var h2 = "ข้อมูลอุปกรณ์";
      //var stbid = "ทดสอบ";
      //var stbmac = "ทดสอบ";
      //var ipaddress = "ทดสอบ";
      //var h3 = "สถานะการเปิดบริการ";
      //var rstatus = "ทดสอบ";
      //var regisdate ="ทดสอบ";
      //var provisdate ="ทดสอบ";
      //var h4 = "การลงทะเบียนบริการเสริม";
      //var hbocpid = "ทดสอบ";
      //var monomaxcpid = "ทดสอบ";
      
      
      var mainheader = "";
      var h1 = "ข้อมูลลูกค้า";
      var account1 = sheetin.getRange(row, 3).getValue();
      var customer1 = sheetin.getRange(row, 4).getValue();
      var package1 = sheetin.getRange(row, 5).getValue();
      var loginid1 = sheetin.getRange(row, 6).getValue();
      var h2 = "ข้อมูลอุปกรณ์ 3BB GIGATV Box ";
      var stbmodel = sheetin.getRange(row, 7).getValue();
      var stbid = sheetin.getRange(row, 8).getValue();
      var stbmac = sheetin.getRange(row, 9).getValue();
      var ipaddress = sheetin.getRange(row, 10).getValue();
      var h3 = "สถานะการเปิดบริการ";
      var rstatus = sheetin.getRange(row, 13).getValue();
      var inregisdate =sheetin.getRange(row, 12).getValue();
      var inprovisdate =sheetin.getRange(row, 11).getValue();
      var h4 = "ข้อมูล services id บริการต่างๆ";
      var stbid = sheetin.getRange(row, 8).getValue();
      var hbocpid = sheetin.getRange(row, 14).getValue();
      var monomaxcpid = sheetin.getRange(row, 15).getValue();
      var regisdate = Utilities.formatDate(new Date(inregisdate),"GMT+7","dd/MM/yyy");
      var provisdate = Utilities.formatDate(new Date(inprovisdate),"GMT+7","dd/MM/yyy");

      var replyccm = {
                 "fulfillmentMessages": [
                  {
                    "platform": "line",
                    "type": 4,
                    "payload": {
                        "line":  {  
                          "type": "flex",
                          "altText": "this is a flex message",
                          "contents": 
                        {
                          "type": "bubble",
                          "size": "giga",
                          "body": {
                            "type": "box",
                            "layout": "vertical",
                            "contents": [
                              {
                                "type": "text",
                                "text": ":: ข้อมูลบริการ :: "+ stbmac,
                                "weight": "bold",
                                 "size": "md",
                                "align": "start",
                                "color": "#0000ff"
                              },
                              {
                                "type": "text",
                                "text":  h1,
                                "weight": "bold",
                                "size": "sm",
                                "align": "center",
                                "color": "#ff0000"
                              },
                              {
                                "type": "text",
                                 "size": "sm",
                                "text": "account no : "+ account1
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "ชื่อลูกค้า : "+ customer1
                              },
                              {
                                "type": "text",
                                "size": "xs",
                                "text": "Package : "+ package1
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "ชื่อ login : "+ loginid1
                              },
                              {
                                "type": "text",
                                "text":  h2,
                                "weight": "bold",
                                "size": "sm",
                                "align": "center",
                                "color": "#ff0000"
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "STB Model : "+ stbmodel
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "STB MAC : "+ stbmac
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "IP Address : "+ ipaddress
                              },
                              {
                                "type": "text",
                                "text":  h3,
                                "weight": "bold",
                                "size": "sm",
                                "align": "center",
                                "color": "#ff0000"
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "สถานะการลงทะเบียน : "+ rstatus
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "วันลงทะเบียน : "+ regisdate
                              },
                              {
                                "type": "text",
                                "size": "sm",
                                "text": "วันติดตั้ง STB : "+ provisdate
                              },
                               {
                                "type": "text",
                                "text":  h4,
                                "weight": "bold",
                                "size": "sm",
                                "align": "center",
                                "color": "#ff0000"
                              },
                               {
                                "type": "text",
                                "size": "xxs",
                                "text": "STB ID : "+ stbid
                              },
                              {
                                "type": "text",
                                "size": "xxs",
                                "text": "HBO CPID : "+ hbocpid
                              },
                              {
                                "type": "text",
                                "size": "xxs",
                                "text": "MONO MAX : "+ monomaxcpid
                              }
                                      ]
                                    }
                                  }
                                }
                              }
                            }]
                            }
 
    return replyccm;
  }

function replayform(sheetin,sheetinname,row)
{
      if (sheetinname == "stberror") 
      { 
      var replymsg ={
                "fulfillmentMessages": [
                  {
                    "platform": "line",
                    "type": 4,
                    "payload": {
                        "line": {
                            "type": "text",
                            "text": sheetin.getRange(row, 2).getValue()
                        }
                    }
                }]
            } 
      }
      else if (sheetinname == "install") 
      { 
      var replymsg ={
                "fulfillmentMessages": [
                  {
                    "platform": "line",
                    "type": 4,
                    "payload": {
                        "line": {
                            "type": "text",
                            "text": sheetin.getRange(row, 2).getValue()
                        }
                    }
                }]
            } 
      }
      else if (sheetinname == "contents")
      { 
      var channelDetail = sheetin.getRange(row, 4).getValue()+"  หมวด : "+ sheetin.getRange(row, 7).getValue()
      var text1 = "3BB GIGATV ช่อง : "+ sheetin.getRange(row, 3).getValue();
      var epg = "epg"+sheetin.getRange(row, 4).getValue()
      var replymsg ={
                "fulfillmentMessages": [
                  {
                    "platform": "line",
                    "type": 4,
                    "payload": {
                        "line": {
                          "type": "template",
                          "altText": "this is a buttons template",
                          "template": {
                            "actions": [
                              {
                                "label": sheetin.getRange(row, 9).getValue(),
                                "text": "status",
                                "type": "message"
                              },
                              {
                                "label": "ผังรายการ",
                                "text": epg,
                                "type": "message"
                              }
                            ],
                            "text": text1,
                            "type": "buttons",
                            "imageAspectRatio": "rectangle",
                            "thumbnailImageUrl": sheetin.getRange(row, 6).getValue(),
                            "title": channelDetail
                          }
                        }
                      }
                    }
                ]
              }
      } 
      else if (sheetinname == "checkccm") 
      { 
      var replymsg = showccm(sheetin,row);
      
      }
      else
      { 
      //var replymsg = showsystest(sheetin,row);
      var replymsg = showsystemstatus(sheetin,row);
      }  

          
    return replymsg;
}


      











