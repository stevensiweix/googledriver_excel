
//----------------------------------------------------------- 啟動

function list_all_files(){

  var folder_id=[];

  folder_id=[
    '1gqkfzdjf85Y9PBD2VQtrTmnu74z2y5qi',
    '1opTVJxb7ktZQByIAVmnml0AuWl5BUDJ2'
  ];

  for(i=0;i<=folder_id.length;i++){
    list_all_files_main(i,folder_id[i]);
    list_user(folder_id[i]);
  }

}

//---------------------------------------------------------- 使用者權限

function list_user(folder_id){

  var folder       = DriveApp.getFolderById(folder_id);

  var Viewers = [];
  Viewers=[
    'xxx1@gmail.com',
    'xxx2@gmail.com'
  ];

  var list_view =folder.getViewers();
  var myArray = [];

  for (i=0;i< list_view.length ;i++){
    //取得email
    check_view = list_view[i].getEmail();
    //丟入陣列
    myArray[i]=check_view;

  }

  //比對含式
  Array.prototype.diff = function (arr) {
    var mergedArr = this.concat(arr);
    return mergedArr.filter(function (e) {
      return mergedArr.indexOf(e) === mergedArr.lastIndexOf(e);
    });
  };

  var diff = Viewers.diff(myArray);

  //比對存入
  for (k=0;k<=diff.length ;k++){
    if(diff[k]!=undefined){
     folder.addViewer(diff[k]);
    }
  }

  /*
  //取消權限
  for (k=0;k<=myArray.length ;k++){
    folder.removeViewer(myArray[k]);
  }
 */


}

//---------------------------------------------------------- 讀取檔案上的資料

function list_all_files_main(num,folder_id){

  //主要函數
  var main = SpreadsheetApp.getActiveSpreadsheet();
  var main = main.getSheets()[num];

  //清空資料表內容
  main.clear();

  //塞入第一行
  main.appendRow(['資料夾','檔案名稱','檔案大小','檔案類型','檔案連結','上傳檔案日期']);
  var range = main.getRange("A1:F1");
  range.setBackground("#ff7e79");

  var folder       = DriveApp.getFolderById(folder_id);
  var folder_child = folder.getFolders();

  //讀取清單
  folder_file_list(main,folder,folder_child);

  //自動排列表格寬度
  Auto_Resize_Column();

  main.setFrozenRows(1);

  //排序為上傳時間
  main.sort(6);

}

//-----------------------------------------------------------檔案清單的抓取

function folder_file_list(main,folder,folder_child){
  //取得第一層資料
  folder_file(main,folder);

  //取得下層資料
  listFolders(main,folder,folder_child);

}

function listFolders(main,folder,childFolders) {
  while(childFolders.hasNext()) {
    var child = childFolders.next();
    folder_file(main,child);
  }
}


function folder_file(main,folder)
{

  //取得資料夾檔案
  var files = folder.getFiles();
  var folder_name = folder.getName();

  //抓取檔案清單
  while (files.hasNext()){

    //取得單一檔案
    file = files.next();

    //取得檔案大小，並格式化
    get_file_size = getReadableFileSizeString(file.getSize());

    data = [
      folder_name,
      file.getName(),
      file.getMimeType(),
      get_file_size,
      file.getUrl(),
      file.getLastUpdated()
    ];

    //寫入檔案
    main.appendRow(data);

  }

}

//-----------------------------------------------------------自動排列表格寬度

function Auto_Resize_Column(){

  var main = SpreadsheetApp.getActiveSpreadsheet();

  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  for (var i = 0 ; i < sheets.length ; i++) {

    var sheet = main.getSheets()[i]; // 第幾個工作表

    //先抓取10個欄位
    for(var x=1; x<=6 ;x++){

      sheet.autoResizeColumn(x);

    }

  }

}

//-----------------------------------------------------------檔案大小的文字讀取

function getReadableFileSizeString(fileSizeInBytes) {
    var i = -1;

    var byteUnits = [' kB', ' MB', ' GB', ' TB', 'PB', 'EB', 'ZB', 'YB'];
    do {
        fileSizeInBytes = fileSizeInBytes / 1024;
        i++;
    } while (fileSizeInBytes > 1024);

    return Math.max(fileSizeInBytes, 0.1).toFixed(1) + byteUnits[i];
};
