function main() {
 // 定期観察
 var presentation = SlidesApp.getActivePresentation();
 const slide = presentation.getSlides();
 const root_date = new Date();
 var date = formatDate(root_date, "MM/dd  (day)  $");
 var text = "定期観察    " + date;
 var iIMG = edit_text(slide, /定期観察\s*\n/, text, 1);
 var fIMG = 0;


 if (iIMG != null) {
   print(iIMG);
   fIMG = replace_IMG(slide, iIMG);
 }
 var re = edit_slide(slide);


 print(slide.length);
 if (iIMG == slide.length - 2) {
   presentation.appendSlide(slide[slide.length - 1]);
 }
  print("----------sumory----------");
 if (iIMG != null) {
   print("Day is written at slide " + iIMG);
   if (fIMG) {
     print("Photo is replaced too at slide " + iIMG);
   } else {
     print("Today's photo wasn't found.");
   }
 } else {
   print("Couldn't find \"定期観察\" in the slide");
 }
   if (re[0] != null) {
     print("Weather was written at these");
     for (var i = 0; i < re.length; i++) {
       print("at slide " + (re[i][0] + 1) + " with date " + re[i][1] + "/" + re[i][2]);
     }
   } else {
     print("Weather isn't written");
   }
  }




function edit_slide(slides) {
 print("edit_slide");
 var root_date = new Date();
 var month = formatDate(root_date, "MM");
 month = parseInt(month);
 var int_day = formatDate(root_date, "dd");
 int_day = parseInt(int_day);
 var weathr_past = 0;
 var weathr_now = get_weather(month);
 var target = /(?<=定期観察\s*).*\$/;
 var index = 0;
 var re = [];
 print("---------search----------");
 for (var i = 0; i < slides.length; i++) {
   const shape = slides[i].getShapes();
   if (shape.length == 0) {continue}
   var textRange = shape[0].getText();
   var text = textRange.asString();
   var matched = text.match(target);


   if (matched == null) {continue}
    
   day = matched[0].match(/(?<=\s*\d+\/)\d+(?=.*\$)/)[0];
   mon = matched[0].match(/(?<=\s*)\d+(?=\/)/)[0];


   if (day == int_day){continue}
     if (mon == month) {
       var weather = weathr_now[day - 1];
     } else {
       if (weathr_past == 0) {
         weathr_past = get_weather(month - 1);
       }
       var weather = weathr_past[day - 1];
     }
      
       var input_text = text.replace(/(?<=.*\)).*$/, (" " + weather[2]));
       textRange.setText(input_text);
       print("Write weather (" + input_text + ") at slide " + i);


 // Second process
 var shape2 = slides[1].getShapes();
 var textRange2 = shape2[1].getText();
 var text2 = textRange2.asString();
 var h_tenp = weather[0];
 var l_tenp = weather[1];


 var grass_h = text2.match(/(?<=:\s*)\d{1,2}(?=cm)/);
 var num_leaf = text2.match(/(?<=:\s*)\d{1,2}(?=枚)/);
 var num_flower = text2.match(/(?<=:\s*)\d{1,2}(?=個)/);
 var plant_info = [grass_h,num_leaf,l_tenp,num_flower];


 for(i=0;i<plant_info.length;i++){
   plant_info[i] = plant_info[i][0]
   if(plant_info[i].length == 1){
     plant_info[i] = ":  " + plant_info[i]
   }else{
     plant_info[i] = ":" + plant_info[i]
   }
 }




 var input_text2 =
   "草丈: " + plant_info[0] + "cm 最高気温    : " + h_tenp+
   "°C\n葉の枚数: " + plant_info[1] + "枚 最低気温.    :"  + l_tenp+
   "°C\n実（花）: " + plant_info[2] + "個  水分含有量:\n" ;


 textRange2.setText(input_text2);
 print("Tenps are set at slide " + 1);


  re[index] = [1, mon, day];
 index++;
 }
 return re;
}




function replace_IMG(slide, iIMG) {
 print("replace_IMG");
 let [imgid, imgname] = find_IMG();


 if (imgname != null) {
   var img = DriveApp.getFileById(imgid);
   img.setTrashed(true);
   print(imgname);


   var images = slide[iIMG].getImages();  // Get all images in the slide
   if (images.length > 0) {
     var image = images[0];  // Assuming you want to replace the first image
     image.replace(img);
     return 1;
   } else {
     print("No images found in the specified slide.");
     return 0;
   }
 } else {
   return 0;
 }
}




function find_IMG() {
 print("---------Drive---------");
 var myDrive = DriveApp.getRootFolder();
 var files = myDrive.getFiles();
 var date = new Date();
 date = formatDate(date, "yyyy/MM/dd");
 var imgFound = false;


 while (files.hasNext()) {
   var file = files.next();
   var mimeType = file.getMimeType();


   // 画像でないファイルは処理しない
   if (!mimeType.match(/image/)) {
     continue;
   }


   var file_name = file.getName();
   var file_id = file.getId();
   var update = file.getLastUpdated();
   update = formatDate(update, "yyyy/MM/dd");


   if (update == date) {
     print("Today's image found: " + file_name);
     imgFound = true;
     return [file_id, file_name];
   }
 }


 if (!imgFound) {
   print("No image found for today");
   return [null, null];
 }
}


function get_weather(date) {
 Logger.log("Fetching weather data...");
 const url = "https://www.data.jma.go.jp/obd/stats/etrn/view/daily_s1.php?prec_no=45&block_no=47682&year=2023&month=" + date;


 const response = UrlFetchApp.fetch(url);
 const content = response.getContentText('UTF-8');
 let get = content.match(/<tr.+?tr>/g);
 let val = [];
  get = get.slice(9);
 for (var i = 0; i < get.length; i++) {
   get[i] = get[i].match(/>.+?<\/td>/g);
   for (var j = 0; j < get[i].length; j++) {
     get[i][j] = get[i][j].slice(1,-5);
     if (get[i][j] == "--") {
       get[i][j] = 0;
     }
   }


   val[i] = get[i].slice(3,-1);
   if (val[i][14] <= 5) {
     val[i][16] = "曇り";
   } 
   if (val[i][13] >= 5) {
     val[i][16] = "晴れ";
   }
   if (val[i][0] >= 30) {
     val[i][16] = "雨";
 }
 }
  var weather = [];


 for (var i = 0; i < val.length; i++) {
   weather[i] = [];
   weather[i][0] = val[i][4]; // 平均気温
   weather[i][1] = val[i][5]; // 最高気温
   weather[i][2] = val[i][16];   // 天候情報
 }
 return weather;
}










function edit_text(slide, target, replace, first) {
 for (var i = 0; i < slide.length; i++) {
   var shape = slide[i].getShapes();
   for (var j = 0; j < shape.length; j++) {
     var textRange = shape[j].getText();
     var text = textRange.asString();
     var matched = text.match(target);


     if (matched != null) {
       textRange.setText(replace);
       if (first == 1) {
         print("Replaced once.");
         return 1;
       }
       print("Replaced all.");
       print("\n" + text);
     }
   }
 }
 print("No target found in the slide");
 return null;
}




function formatDate(date, format) {
 format = format.replace(/yyyy/g, date.getFullYear());
 format = format.replace(/MM/g, ((date.getMonth() + 1)));
 format = format.replace(/dd/g, (date.getDate()));
 format = format.replace(/HH/g, (date.getHours()));
 format = format.replace(/mm/g, (date.getMinutes()));
 format = format.replace(/ss/g, (date.getSeconds()));


 let week_day = ["日", "月", "火", "水", "木", "金", "土"];
 var day = week_day[date.getDay()];
 format = format.replace(/day/g, day);
 return format;
}


function print(text) {
 console.log(text);
}

