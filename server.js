function myFunction() {
  const now=new Date();
  Logger.log(now);
  Logger.log(now.getTime());
  const thirtyMinutesAgo=new Date(now.getTime()-60*60*1000);
  Logger.log(thirtyMinutesAgo);

  const query="(from:students.cdc.2026@vitap.ac.in OR from:jaladiyaswanth2005@gmail.com) is:unread";


  const threads=GmailApp.search(query);
  const recentThreads= threads.filter(thread=> thread.getLastMessageDate()>= thirtyMinutesAgo);

  if(recentThreads.length===0) return;
  Logger.log(recentThreads);

  recentThreads.forEach(thread=>{
    const messages=thread.getMessages();
    messages.forEach(msg=>{
      const subject=msg.getSubject();

      if(!subject.toLowerCase().includes("test")){
         Logger.log("No test ");
         return;}

      const attchments=msg.getAttachments();
      attchments.forEach(attchment=>{
        if(attchment.getContentType().includes("sheet")|| attchment.getName().match(/\.(xlsx|xls)$/)){
          const blob=attchment.copyBlob();
          const file=DriveApp.createFile(blob);
          const sheetfile=SpreadsheetApp.openById(file.getId());

          let found=false;
          sheetfile.getSheets().forEach(sheet=>{
            const data=sheet.getDataRange().getValues();

            data.forEach(row=>{
              if(row.join(' ').includes(my_name) || row.join(' ').includes(my_reg_number)){
                  found=true;
                  
              }
            })
          });

          if(found){
            let venue="";
            if(msg.getSubject().toLowerCase().includes("own location")){
              venue="Own Location";
            }

            else venue=extractvenue(attchments)||"TBD";
          
            const start=extractDateTimeFromText(msg.getSubject());
      const end=new Date(start.getTime()+60*60*1000);
              CalendarApp.createEvent(
            msg.getSubject(),
            start,
            end,
            {
              location:`${venue}`,
              description:msg.getThread().getPermalink()+
              "\nsee this"

            }
            
            
          );
            Logger.log("Created from if");
          }
        }
        
      });
      let venue="";
            if(msg.getSubject().toLowerCase().includes("own location")){
            
            venue="Own location";}

      const start=extractDateTimeFromText(msg.getSubject());
      const end=new Date(start.getTime()+60*60*1000);
      CalendarApp.createEvent(
            msg.getSubject(),
            start,
            end,
            {
              location:`${venue}`,
              description:msg.getThread().getPermalink()+
              "\nsee this"

            }
            
            
          );
          Logger.log("Created from else");
      
        


    })
  })
}


function extractDateTimeFromText(text){
  try{
    Logger.log(`Extracting date/time from: ${text}`);
  

  const extractedDate=extractDate(text);
  const extractedTime=extractTime(text);

  if(extractedDate && extractedTime){
    const datetimestr=`${extractedDate} ${extractedTime}`;
  const parsedDate=new Date(datetimestr);
  return parsedDate;
  }

  if(extractedDate && !extractedTime){
    const datatimestr=`${extractedDate} 10:00 AM`;
    const parsedDate=new Date(datatimestr);
    return parsedDate;
  } 
  } catch(error){
    Logger.log("Error parsing Date/Time"+error.toString());
}}


function extractDate(text){
  
        try{
         let  monthMap = {
            'jan': 'January', 'feb': 'February', 'mar': 'March', 'apr': 'April',
            'may': 'May', 'jun': 'June', 'jul': 'July', 'aug': 'August',
            'sep': 'September', 'oct': 'October', 'nov': 'November', 'dec': 'December'
          };
          Logger.log("extrcating from text");
          const patterns = [
  /on\s+(\d{1,2})\.(\d{1,2})\.(\d{4})/i,                         // on 08.09.2025
  /on\s+(\d{1,2})(?:st|nd|rd|th)?\s+(\w+)\s+(\d{4})/i,           // on 8th September 2025
  /on\s+(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3})\s+(\d{4})/i,   // on 8th Sep 2025
  /on\s+(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/i                  // on 08/09/2025 or on 08-09-2025
];

    for(let i=0;i<patterns.length;i++){
      let match= text.match(patterns[i]);
      if(match){
        let day,month,year,formattedDate;

        if(i==0 || i==3){
          day=match[1];
          month=match[2];
          year=match[3];
          formattedDate=`${month}/${day}/${year}`;
        }
        else if(i==1){
           day=match[1];
          month=match[2];
          year=match[3];
          formattedDate=`${month} ${day} ${year}`;
        }
        else if(i==2){
          day=match[1];
          const short=match[2];
          year=match[3];

          const fullmonth=monthMap[short.toLowerCase()];
          if(!fullmonth) continue;
          formattedDate=`${fullmonth} ${day} ${year}`;
        }
        Logger.log("Extracted Date:"+formattedDate);
        const finalDate=new Date(formattedDate);
        if(!isNaN(finalDate.getTime())){
          return formattedDate;
        }
        
      }
    }

        }catch(error){
          Logger.log("Error extracting");
          return null;
        }  
}



function extractTime(text){
  try{
    Logger.log(`Extracting time from :${text}`);

    const pattern1= /(\d{1,2})[:.](\d{2})\s*([ap])\.?m\.?/i;

    let match=text.match(pattern1);

    if(match){
      const hour=match[1];
      const minute=match[2];
      const ampm=match[3].toUpperCase()+'M';
      const time=`${hour}:${minute}:${ampm}`;
      Logger.log("found pattern1" );
      return time;
    }
  const pattern2 = /by\s+(\d{1,2})[:.](\d{2})\s*(am|pm)/i;
    match = text.match(pattern2);

    if (match) {
      const hour = match[1];
      const minute = match[2];
      const ampm = match[3].toUpperCase();
      const time = `${hour}:${minute} ${ampm}`;
      Logger.log(`Found time pattern 2: ${time}`);
      return time;
    }

  }catch(error){
    Logger.log("error extrcating time"+error.toString());
    return null;
  }

}



function extractvenue(attchments){
  try{
    Logger.log("Processing venue list");
    let extractedVenue=null;
    attchments.forEach(attchement=>{
      const fileName=attchment.getName();
      if(fileName.toLowerCase().includes('venue')){
        Logger.log('Found venues list');
        extractedVenue=extrcatfromlist(attchement,name,re);
        if(extractedVenue){
          Logger.log("found venue in list");
          return extractedVenue;
        }
      }
    })
    return null;
  }catch(error){
    Logger.log("Error in venue");
    return;
  }
}


function extrcatfromlist(attchement,name,reg){
return null;
}
