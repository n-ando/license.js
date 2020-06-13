//
// Windows および Microsoft製品ライセンス情報収集スクリプト
//                      by Noriaki Ando <n-ando@aist.go.jp>
//                      date: 2010.01.06
//
// このスクリプトはWindowsおよびMicrosoft製品のライセンス情報を収集し
// csv ファイルに出力します。すべてのMicrosoft製品に対応しているわけで
// はありません。
//
var ProductKey = function ( ) {
  this.office_regkeys = {
      // 参考 http://d.hatena.ne.jp/frontline/20081222/p1
      "Office XP Professional":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\10.0\\Registration\\{90280411-6000-11D3-8CFE-0050048383C9}",
      "Office XP Professional(English)":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\10.0\\Registration\\{90110409-6000-11D3-8CFE-0050048383C9}",
      "Office 2003 Standard/Professional":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\11.0\\Registration\\{90110411-6000-11D3-8CFE-0150048383C9}",
      "Visio 2003 Standard/Professional":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\11.0\\Registration\\{90510411-6000-11D3-8CFE-0150048383C9}",
      "Office 2007 Professional":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\12.0\\Registration\\{90120000-0014-0000-0000-0000000FF1CE}",
      "Office 2007 Professional Plus":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\12.0\\Registration\\{90120000-0011-0000-0000-0000000FF1CE}",
      "Office 2007 Enterprise":"HKEY_LOCAL_MACHINE\\SOFTWAWRE\\Microsoft\\Office\\12.0\\Registration\\{90120000-0030-0000-0000-0000000FF1CE}",
      "Office 2007 Ultimate":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\12.0\\Registration\\{91120000-002E-0000-0000-0000000FF1CE}",
      "Visio 2007 Professional":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\12.0\\Registration\\{90120000-0051-0000-0000-0000000FF1CE}",
      "Project 2007 Professional":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\12.0\\Registration\\{90120000-003B-0000-0000-0000000FF1CE}",
      "SQL Server 2005":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Microsoft SQL Server\\90\\ProductID",
      "Office 2007 Personal":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\12.0\\Registration\\{91120000-0018-0000-0000-0000000FF1CE}",
      "Office PowerPoint 2007":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\12.0\\Registration\\{91120000-0018-0000-0000-0000000FF1CE}",
      "Office 2010 Professional Plus":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\14.0\\Registration\\{90140000-002A-0000-1000-0000000FF1CE}",
      "Office 2010 Professional Plus on x64":"HKEY_LOCAL_MACHINE\\SOFTWARE\\Wow6432Node\\Microsoft\\Office\\14.0\\Registration\\{91140000-0011-0000-0000-0000000FF1CE}"
  }

  this.vs_regkeys = {
    "Visual Studio 2005 Professional": "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\VisualStudio\\8.0\\Registration\\2000.0x0000"
  }


  this.WshShell = new ActiveXObject( "WScript.Shell" );
  this.read_reg = function ( reg_key_name ){
    return  this.WshShell.RegRead(reg_key_name);
  }
  this.decodeDigitalProductId = function( reg_binary_value ){
    reg_binary_value = reg_binary_value.toArray();
    reg_binary_value = reg_binary_value.slice( 52, 67 );
    decodedChars = Array(24);
    digits = Array("B","C","D","F","G","H","J","K","M","P","Q","R","T","V","W","X","Y","2","3","4","6","7","8","9");
    for (var i=1; i < 30; i++)
      {
        if((i % 6) == 0)
          {
            decodedChars.push("-");
          }
        else
          {
            k=0;
            for (var j = 14; j >= 0; j--)
              {
                k = ( k << 8 ) ^ reg_binary_value[j];
                reg_binary_value[j] =  parseInt( k /digits.length );
                k = k % digits.length;
              }
            decodedChars.push(digits[k]);
          }
      }
    var aa = decodedChars.reverse().join('').split("-")
    //    aa.pop();
    //    aa.push("*****")
    return aa.join("-");
  }
  
  // office() function
  this.office = function(){
    var result = [];
    for ( key_name in this.office_regkeys ){
      try{
        var productkey = this.read_reg(this.office_regkeys[key_name] + '\\DigitalProductId');
        var productid  = this.read_reg(this.office_regkeys[key_name] + '\\ProductId');
        result.push( key_name + ":" + this.decodeDigitalProductId( productkey ) + ":" + productid);
      } catch(e) {
        //
      }
    }
    return result
  }
  
  // windows() functions
  this.windows = function (){
    return this.WindowsName() + ":" + this.Windows_ProductKey() + ":" + this.Windows_ProductId();
  }

  this.Windows_ProductKey = function (){
    var ret0 = this.read_reg('HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\DigitalProductId');
    WScript.Echo("DefaultProductKey\\DigitalProductId" + ret0);
    var ret1 = this.read_reg('HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\DefaultProductKey\\DigitalProductId');
    WScript.Echo("DefaultProductKey\\DigitalProductId" + ret1);
    return this.decodeDigitalProductId( ret1 );
  }

  this.Windows_ProductId = function () {
    var ret0 = this.read_reg('HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\DefaultProductKey\\ProductId');
    WScript.Echo("DefaultProductKey\\DigitalProductId" + ret0);
    var ret1 = this.read_reg('HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProductId');
    WScript.Echo("DefaultProductKey\\DigitalProductId" + ret1);
    return ret1
  }

  this.WindowsName = function(){
    var name = "";
    var pack = "";
    try {
      name = this.read_reg('HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProductName');
      pack = this.read_reg('HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\CSDVersion');
    } catch (e) {
      //
    }
    return name + "("+pack+")";
  }

  this.visualstudio = function () {
    var result = [];
    for ( key_name in this.vs_regkeys ){
      try{
        var productkey = this.read_reg(this.vs_regkeys[key_name] + '\\PIDKEY');
        var productid  = this.read_reg(this.vs_regkeys[key_name] + '\\ProductID');
        result.push( key_name + ":" + productkey + ":" + productid);
      } catch(e) {
        //
      }
    }
    return result
  }

}
  
function get_nic_info()
{
  WshShell = new ActiveXObject( "WScript.Shell" );
  var objExec = WshShell.Exec("ipconfig.exe /all");
  var entries = Array();
  var nicentry = Array();
  while (!objExec.StdOut.AtEndOfStream)
  {
    line = objExec.StdOut.ReadLine();
    line = line.replace("\r", "");
    if (line.match(/Description|説明/)) {
      var ss = line.split(":");
      nicentry["nic"] = ss[1];
    }
    //    if (line.match(/Physical Address/)) {
    if (line.match(/[0-9A-Fa-f][0-9A-Fa-f]-[0-9A-Fa-f][0-9A-Fa-f]-[0-9A-Fa-f][0-9A-Fa-f]-[0-9A-Fa-f][0-9A-Fa-f]-[0-9A-Fa-f][0-9A-Fa-f]-[0-9A-Fa-f][0-9A-Fa-f]/)) {
      var ss = line.split(":");
      nicentry["mac"] = ss[1];
    }
    if (line.match(/IP Address|IPv4/)) {
      var ss = line.split(":")
      nicentry["ip"] = ss[1];
      entries.push(nicentry);
    }
  }
  return entries;
}

//
// return: 
//
//
function get_vender_info()
{
  var wmi = GetObject("winmgmts:\\\\.\\root\\cimv2");
  var items = wmi.ExecQuery("Select * from Win32_ComputerSystemProduct");
  var e = new Enumerator(items);
  var item = e.item();

  var prop = new Array();
  prop["product"] = item.Name;
  prop["vendor"]  = item.Vendor;
  prop["serialid"] = item.IdentifyingNumber;

  var sysitems = wmi.ExecQuery("Select * from Win32_ComputerSystem");
  var syse = new Enumerator(sysitems);
  var sysitem = syse.item();
  prop["username"] = sysitem.UserName.split("\\")[1];
  prop["hostname"] = sysitem.UserName.split("\\")[0];
  return prop;
}


function licenselist_wiki() {
  var fso = new ActiveXObject( "Scripting.FileSystemObject" );
  lfile = fso.OpenTextFile("license.txt", 8, true);

  var text = String();
  var p = new ProductKey();

  text += "| 使用者      |              |\n";
  text += "| 部屋番号    |              |\n";
  text += "| 用途        | NEDO知能化PJ |\n";

  var vinfo = get_vender_info();
  text += "|機種名       |" + vinfo["product"]  + "|\n";
  text += "|ベンダ名     |" + vinfo["vendor"]   + "|\n";
  text += "|シリアル番号 |" + vinfo["serialid"] + "|\n";

  text += "|>|ソフトウエアプロダクトキー|\n"

  var tmp = p.windows().split(":");
  text += "|" + tmp[0] + "|" + tmp[1] + "|\n";

  var offices = p.office();
  for (var i = 0; i < offices.length; i++)
    {
      var otmp = offices[i].split(":");
      text += "|" + otmp[0] + "|" + otmp[1] + "|" + otmp[2] + "|\n";
    }

  var vs = p.visualstudio();
  for (var i = 0; i < vs.length; i++)
    {
      var vtmp = vs[i].split(":");
      text += "|" + vtmp[0] + "|" + vtmp[1] + "|" + vtmp[2] + "|\n";
    }
  //WScript.Echo(text);
  lfile.Write(text);
  lfile.Close();
}

function es(str)
{
  if (str.match(",")) {
    return "\"" + str + "\"";
  }
  return str;
}

function licenselist_csv()
{
  
  var vinfo = get_vender_info();
  var nicinfo = get_nic_info();
  var text = String();
  var fname = "license_" + vinfo["username"] + "_" + vinfo["hostname"] + ".csv";
  var fso = new ActiveXObject( "Scripting.FileSystemObject" );
  lfile = fso.OpenTextFile(fname, 2, true);



  lfile.Write("ソフトウエア製品名, ハードウエアシリアル番号, 資産番号, ユーザ名, ホスト名, メーカー名, 機種名, MACアドレス, プロダクトID, バージョン, ライセンスキー\n");
  var pcinfo = String();
  // ハードウエアシリアル番号, 資産番号, ユーザ名, ホスト名, メーカー名, 機種名, MACアドレス
  pcinfo  = es(vinfo["serialid"]) + ",,";
  pcinfo += es(vinfo["username"]) + ",";
  pcinfo += es(vinfo["hostname"]) + ",";
  pcinfo += es(vinfo["vendor"]  ) + ",";
  pcinfo += es(vinfo["product"] ) + ",";
  pcinfo += es(nicinfo[0]["mac"]);
  
  // Windows product id
  var p = new ProductKey();
  var wtmp = p.windows().split(":");
  text = es(wtmp[0]) + "," + pcinfo + "," + es(wtmp[2]) + ",," + es(wtmp[1]) + "\n";
  lfile.Write(text);
  
  // Office product id
  var offices = p.office();
  for (var i = 0; i < offices.length; i++)
    {
      text = "";
      var otmp = offices[i].split(":");
      text = es(otmp[0]) + "," + pcinfo + "," + es(otmp[2]) + ",," + es(otmp[1]) + "\n";
      lfile.Write(text);
    }
  
  var vs = p.visualstudio();
  for (var i = 0; i < vs.length; i++)
    {
      text = "";
      var vtmp = vs[i].split(":");
      text = es(vtmp[0]) + "," + pcinfo + "," + es(vtmp[2]) + ",," + es(vtmp[1]) + "'\n";
      lfile.Write(text);
    }
  lfile.Close();
  WScript.Echo("ライセンス情報ファイルを作成しました: " + fname);
}

//licenselist_wiki();
licenselist_csv();
