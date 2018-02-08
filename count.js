#title = "CharRank"
#tooltip = "文字出現量ランキング"
/*------------------------------------------------------------
CharRank -文字出現量ランキング- var 0.91
 テキスト内のそれぞれのキャラクタの個数を数え、ランキング形式で表示。
 英文の中で最も出現量が多いのがeと言うのを検証したりするときなんかに使ってみてください。
    ※サロゲートペア対応
 
    テキストエディタMery用マクロ
    By 群生系パセリ (2015/12/27)
------------------------------------------------------------*/
var $d = Document;
var ActText = $d.Text;
var ActSize = ActText.length;
var ActSizeS = ActSize - (ActText.match(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g)||[]).length;
var ActPath = $d.FullName;
var Size = new Object();
var CaraSort = new Array();
 
//新規ファイルを作成
Editor.NewFile();
var NewDoc = Editor.Documents.Item(Editor.Documents.Count-1);
NewDoc.Write("Loading...");
 
 
var DText = "CharRank -文字出現量ランキング- var 0.90\n\tPath\t: "+ActPath+"\n\tLength\t: "+ActSizeS+"\n-------------";
 
 
//文字数カウント
 
var iCara="",iCarai=0;
var sarohi=0,sarolo=0;
for(var i=0; i<ActSize; i++){
    iCarai=ActText.charAt(i).charCodeAt();
    if(iCarai>=55296 && iCarai<=56319){
        sarohi=iCarai;
    }else{
        if(iCarai>=56320 && iCarai<=57343){
            sarolo=iCarai;
            iCarai = 0x10000 + (sarohi - 0xD800) * 0x400 + (sarolo - 0xDC00)
        }
        iCara="U+"+("0000"+iCarai.toString(16)).slice(-5);
        if(Size[iCara]==void 0){
            Size[iCara]=1;
        }else{
            Size[iCara]++;
        }
    }
}
//ソート
var i=0;
for(var prop in Size){
    CaraSort[i]=prop+" "+Size[prop];
    i++;
}
CaraSort.sort(function(a,b){
    as=parseInt(a.split(" ")[1]);
    bs=parseInt(b.split(" ")[1]);
    return bs-as;
});
//ランキング作成
var iCode="";
var LankZ=0,LankZs=""
var CaraSortlen=CaraSort.length.toString().length
for(var i=0;i<CaraSortlen;i++){
    LankZ++;
    LankZs+="0";
}
var Lanki="",ikazu=0,ikazube;
for(var i=0;i<CaraSort.length;i++){
    iCode=CaraSort[i].split(" ")[0];
    if(iCode.slice(0,3)!="U+0"){
        iCarai=parseInt(iCode.slice(-5),16);
        sarohi=(iCarai - 0x10000) / 0x400 + 0xD800
        sarolo=(iCarai - 0x10000) % 0x400 + 0xDC00
        iCara=String.fromCharCode(sarohi)+String.fromCharCode(sarolo)
    }else{
        iCara=String.fromCharCode(parseInt(iCode.slice(-5),16));
    }
    if(iCara=="\n"){iCara="{LF}"}
    if(iCara=="\t"){iCara="{TAB}"}
    ikazu=parseInt(CaraSort[i].split(" ")[1]);
    if(ikazube!=ikazu){
        Lanki=(LankZs+(i+1).toString()).slice(-LankZ);
    }
    DText+="\n"+Lanki+"\t"+iCara+"\t["+iCode+"]\t"+ikazu;
    ikazube=ikazu;
}
DText=DText.replace(/U\+0/g," U+")
 
//結果を出力
NewDoc.Selection.DeleteLeft(10);
NewDoc.Write(DText);
 
NewDoc.Selection.SetActivePos(0) ;