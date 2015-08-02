/*setcompatmode.js*/
var patharray = new Array(2);
var filecount_num = new Array(2);
for(var i=0;i<2;i++){
  patharray[i]= new Array();
  filecount_num[i]=0;
}
var fso = new ActiveXObject("Scripting.FileSystemObject");

var objApl = WScript.CreateObject("Shell.Application");

var title = "Select the Folder";
var option = 0x0050;
var root = "";
var folder = objApl.BrowseForFolder(0, title, option, root);

if (folder == null) {
    WScript.Echo("You must select the folder");
    WScript.Quit();//quit
}else{

  WScript.echo(folder.Self.Path);
}
var ScriptFolderPath = String(WScript.ScriptFullName).replace(WScript.ScriptName,"");

var WshShell = WScript.CreateObject("WScript.Shell");
var writevalue = "HKCU\\Software\\microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced";

var Origin_FolderOptReg = new Array();
Origin_FolderOptReg[0]=WshShell.RegRead(writevalue+"\\Hidden");//元の状態を一時保存
Origin_FolderOptReg[1]=WshShell.RegRead(writevalue+"\\HideFileExt");

//GetExtensionNameを正常に利用するには拡張子等の表示を有効にする必要あり
WshShell.RegWrite(writevalue+"\\Hidden", 0x01, "REG_DWORD");//隠しファイルとフォルダ表示
WshShell.RegWrite(writevalue+"\\HideFileExt", 0x00, "REG_DWORD");//拡張子表示
WshShell.RegWrite("HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers\\","","REG_SZ");//キー未作成であっても対応できるように該当キーを作成
WshShell.Run( "RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters ",1,true);

var objFolder = objApl.NameSpace(folder.Self.Path);

var objFolderItems = objFolder.Items();

GetFilePathFromExtensionName(objFolderItems,"exe","msi");

objFolderItems = null;
objFolder = null;
objApl = null;

WScript.echo(filecount_num[0]);
var value = "~ WINXPSP3";
if(wmi_addComatinfo(patharray[0],value,filecount_num[0])!=0){
  WScript.echo("error");
}
WScript.echo(filecount_num[1]);
value = "~ MSIAUTO";
if(wmi_addComatinfo(patharray[1],value,filecount_num[1])!=0){
  WScript.echo("error");
}else{
  WScript.echo("End");
}
//WshShell.RegWrite(writevalue+"\\Hidden", Origin_FolderOptReg[0], "REG_DWORD");
//WshShell.RegWrite(writevalue+"\\HideFileExt", Origin_FolderOptReg[1], "REG_DWORD");

function wmi_addComatinfo(add_name ,add_value, f_num){//Nameに\\を含む場合、RegWriteが有効でないためWMIを利用する
  var AppCompatFlagsRegistryKey = "Software\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers";
  var name;
  var Data = add_value;
  var result=0;
  var objRegistry = GetObject("winmgmts://./root/default:StdRegProv");
    for(var i=0;i<f_num;i++){
        name=add_name[i];
      try {
        result = objRegistry.SetStringValue(0x80000001 /*HKCU*/, AppCompatFlagsRegistryKey, name, Data);
      } catch (e) {
        WScript.echo(e.message);
        result++;
     }
    }
  return result;
  }

function GetFilePathFromExtensionName(tmpFolderItems , exn0,exn1) {
    var objFolderItemsB;
    var objItem;

    for (var i=0;i<tmpFolderItems.Count;i++){

        objItem = tmpFolderItems.Item(i);
        if (objItem.IsFolder==true) {

            WScript.echo(objItem.Name);
           objFolderItemsB = objItem.GetFolder;
           GetFilePathFromExtensionName(objFolderItemsB.Items(),exn0,exn1);
        } else {

           if(fso.GetExtensionName(objItem.Name)==exn0||fso.GetExtensionName(objItem.Name)==exn0.toUpperCase()){
            patharray[0][filecount_num[0]]=objItem.Path;
            filecount_num[0]++;
           }else if(fso.GetExtensionName(objItem.Name)==exn1||fso.GetExtensionName(objItem.Name)==exn1.toUpperCase()){
            patharray[1][filecount_num[1]]=objItem.Path;
            WScript.echo(patharray[1][filecount_num[1]]);
            filecount_num[1]++;
           }
        }

    }

    objItem = null;
    objFolderItemsB = null;

}
