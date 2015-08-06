/*setcompatmode.js ver0.0.5*/
/*
setcompatmode.js

Copyright (c) 2015 prprhyt

This software is released under the MIT License.
http://opensource.org/licenses/mit-license.php
*/

var patharray = new Array(2);
for(var i=0;i<2;i++){
  patharray[i]= new Array();
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

var exn_array =["exe","msi"];//検索対象の拡張子
var value_array=["~ WINXPSP3","~ MSIAUTO"];//互換モードの要素

if(exn_array.length!=value_array.length){
    WScript.Echo("Different number of ExtensionName array and Value array");
    WScript.Quit();//quit
}

GetFilePathFromExtensionName(objFolderItems,exn_array);
var files_sum=0;
for(var i=0;i<exn_array.length;i++){
  files_sum += patharray[i].length;
}

if(wmi_addComatinfo(patharray,value_array)!=0){
   WScript.echo("error");
 }else{
   WScript.echo("End:"+folder.Self.Path+"\n"+files_sum+"files complete.");
}

//WshShell.RegWrite(writevalue+"\\Hidden", Origin_FolderOptReg[0], "REG_DWORD");
//WshShell.RegWrite(writevalue+"\\HideFileExt", Origin_FolderOptReg[1], "REG_DWORD");

objFolderItems = null;
objFolder = null;
objApl = null;

function wmi_addComatinfo(add_name ,add_value){//Nameに\\を含む場合、RegWriteが有効でないためWMIを利用する
  var AppCompatFlagsRegistryKey = "Software\\Microsoft\\Windows NT\\CurrentVersion\\AppCompatFlags\\Layers";
  var name;
  var Data;
  var result=0;
  var objRegistry = GetObject("winmgmts://./root/default:StdRegProv");
    for(var i=0;i<add_value.length;i++){
      Data = add_value[i];
      for(var j=0;j<add_name[i].length;j++){
        name=add_name[i][j];
        try {
          result = objRegistry.SetStringValue(0x80000001 /*HKCU*/, AppCompatFlagsRegistryKey, name, Data);
        } catch (e) {
           result++;
        }
      }
    }
  return result;
}

function GetFilePathFromExtensionName(tmpFolderItems , exn_array_s) {
    var objFolderItemsB;
    var objItem;

    for (var i=0;i<tmpFolderItems.Count;i++){

        objItem = tmpFolderItems.Item(i);
        if (objItem.IsFolder==true) {
           objFolderItemsB = objItem.GetFolder;
           GetFilePathFromExtensionName(objFolderItemsB.Items(),exn_array_s);
        } else {
         for(var j=0;j<exn_array_s.length;j++){
           if(fso.GetExtensionName(objItem.Name)==exn_array_s[j].toLowerCase()||fso.GetExtensionName(objItem.Name)==exn_array_s[j].toUpperCase()){
            patharray[j][patharray[j].length]=objItem.Path;
           }
         }
        }

    }

    objItem = null;
    objFolderItemsB = null;

}
