
const my_name=  "Your_Name;
const my_reg_number="Your_Registration_Number";
function myFunction() {
  const now=new Date();
  Logger.log(now);
  Logger.log(now.getTime());
  const thirtyMinutesAgo=new Date(now.getTime()-120*60*1000);
  Logger.log(thirtyMinutesAgo);


  const query="(from:students.cdc2026@vitap.ac.in  from:placement@vitap.ac.in OR to:students.cdc2026@vitap.ac.in) is:unread";


  const threads=GmailApp.search(query);
  const recentThreads= threads.filter(thread=> thread.getLastMessageDate()>= thirtyMinutesAgo);

  if(recentThreads.length===0) return;
  Logger.log(recentThreads);

  recentThreads.forEach(thread=>{
    const messages=thread.getMessages();
    messages.forEach(msg=>{
      const subject=msg.getSubject();

      if(!subject.toLowerCase().includes("talk") && !subject.toLowerCase().includes("test")){
         Logger.log("No test ");
         return;}

      const attachments=msg.getAttachments();
      attachments.forEach(attachment=>{
        if(attachment.getContentType().includes("sheet")|| attachment.getName().match(/\.(xlsx|xls)$/)){
          const blob=attachment.copyBlob();
          // blob.setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
          // const file=DriveApp.createFile(blob);
          // const sheetfile=SpreadsheetApp.openById(file.getId());
          

        //   const fileMetadata = {
        //     'name': attachment.getName().replace(/\.(xlsx|xls)$/, ''),
        //     'mimeType': 'application/vnd.google-apps.spreadsheet'
        // };
        
        // const media = {
        //     mimeType: attachment.getContentType(),
        //     body: blob.getBytes()
        // };
        
        // const file = Drive.Files.create(fileMetadata, media, {
        //     uploadType: 'multipart'
        // });
        
        // const sheetfile = SpreadsheetApp.openById(file.id);

      const fileMetadata = {
            'name': attachment.getName().replace(/\.(xlsx|xls)$/, ''),
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        };
        
        // Use the blob directly, not getBytes()
        const file = Drive.Files.insert(fileMetadata, blob,{convert: true});
        
        const sheetfile = SpreadsheetApp.openById(file.id);

          let found=false;
          // sheetfile.getSheets().forEach(sheet=>{
          //   const data=sheet.getDataRange().getValues();

          //   data.forEach(row=>{
          //     if(row.join(' ').includes(my_name) || row.join(' ').includes(my_reg_number)){
          //         found=true;
                  
          //     }
          //   })
          // });

          for (let sheet of sheetfile.getSheets()) {
  const data = sheet.getDataRange().getValues();
  for (let row of data) {
    if (row.join(' ').includes(my_name) || row.join(' ').includes(my_reg_number)) {
      found = true;
      break; // exit row loop
    }
  }
  if (found) break; // exit sheet loop
}




          DriveApp.getFileById(file.id).setTrashed(true);

          if(found){
            let venue="";
            if(msg.getSubject().toLowerCase().includes("own location")){
              venue="Own Location";
            }

            else venue=extractvenue(attachments)||"TBD";
          
            let start=extractDateTimeFromText(msg.getSubject());


            if (!start) {
              Logger.log("No date found, using tomorrow 10:00 AM as default");
              const tomorrow = new Date();
              tomorrow.setDate(tomorrow.getDate() + 1);
              tomorrow.setHours(10, 0, 0, 0); // 10:00 AM
              start = tomorrow;
            }
            const end=new Date(start.getTime()+60*60*1000);

            //Creating Calendar Event
             try{ CalendarApp.createEvent(
            msg.getSubject(),
            start,
            end,
            {
              location:`${venue}`,
              description:msg.getThread().getPermalink()+
              "\nsee this"

            }
            
            
          );
          
              if(venue==="TBD") Logger.log("Created calendar event with TBD");
            else Logger.log("Created calendar event with location");
          }catch(error){
              Logger.log("Error creating calendar event " +error.toString());
          }
          }
        }
        
      });

      /*
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
      */
        


    })
  })
}


function create_calender(){}


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
    const datetimestr=`${extractedDate} 10:00 AM`;
    const parsedDate=new Date(datetimestr);
    return parsedDate;
  } 

  
  return null;
  } catch(error){
    Logger.log("Error parsing Date/Time " +error.toString());
    return null;
}}


function extractDate(text){
  
        try{
         let  monthMap = {
            'jan': 'January', 'feb': 'February', 'mar': 'March', 'apr': 'April',
            'may': 'May', 'jun': 'June', 'jul': 'July', 'aug': 'August',
            'sep': 'September', 'oct': 'October', 'nov': 'November', 'dec': 'December'
          };
          Logger.log("extracting Date from text");
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


      return null;
        }catch(error){
          Logger.log("Error extracting date");
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
      const time=`${hour}:${minute} ${ampm}`;
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

    return null;

  }catch(error){
    Logger.log("error extracting time:" +error.toString());
    return null;
  }

}



function extractvenue(attachments){
  try{
    Logger.log("Processing venue list");
    let extractedVenue=null;
    for(let attachment of attachments){
      const fileName=attachment.getName();
      if(fileName.toLowerCase().includes('venue')){
        Logger.log('Found venues list');
        extractedVenue=extractfromlist(attachment,my_name,my_reg_number);
        if(extractedVenue){
          Logger.log("found venue in list");
          return extractedVenue;
        }
      }
    }
    return null;
  }catch(error){
    Logger.log("Error in venue");
    return null;
  }
}


// function extractfromlist(attachment,name,reg){
//     const blob=attachment.copyBlob();
//           const file=DriveApp.createFile(blob);
//           const sheetfile=SpreadsheetApp.openById(file.getId());
//     const data=sheetfile.getDataRange().getValues();
//     const venueIndex=data[0].findIndex(h=>h.toString().toLowerCase().includes("venue"));
//     if(venueIndex==-1) return null;

//     for(let i=1;i<data.length;i++){
//       const row=data[i].join(' ').toLowerCase();
//       if(row.includes(name) || row.includes(reg)){
//         DriveApp.getFileById(file.getId()).setTrashed(true);
//         return data[i][venueIndex];
//       }
//     }
//         DriveApp.getFileById(file.getId()).setTrashed(true);

//     return null;
// }


// function extractfromlist(attachment, name, reg){
//   let tempFile = null;
//   try {
//     const blob = attachment.copyBlob();
//     tempFile = DriveApp.createFile(blob);
//     const sheetfile = SpreadsheetApp.openById(tempFile.getId());
//     const data = sheetfile.getDataRange().getValues();
    
//     const venueIndex = data[0].findIndex(h => h.toString().toLowerCase().includes("venue"));
//     if(venueIndex === -1) return null;

//     for(let i = 1; i < data.length; i++){
//       const row = data[i].join(' ').toLowerCase();
//       if(row.includes(name.toLowerCase()) || row.includes(reg.toLowerCase())){
//         return data[i][venueIndex];
//       }
//     }
//     return null;
    
//   } catch(error) {
//     Logger.log("Error in extractfromlist: " + error.toString());
//     return null;
//   } finally {
//     // Clean up temporary file
//     if(tempFile) {
//       try {
//         DriveApp.getFileById(tempFile.getId()).setTrashed(true);
//       } catch(e) {
//         Logger.log("Error deleting temp file: " + e.toString());
//       }
//     }
//   }
// }



function extractfromlist(attachment, name, reg) {
  let tempFile = null;
  try {
    const blob = attachment.copyBlob();

    // Convert uploaded Excel file to Google Sheet
    const fileResource = {
      title: attachment.getName(),
      mimeType: MimeType.GOOGLE_SHEETS
    };
    tempFile = Drive.Files.insert(fileResource, blob, {convert: true});

    const sheetfile = SpreadsheetApp.openById(tempFile.id);
    const data = sheetfile.getDataRange().getValues();

    const venueIndex = data[0].findIndex(h => h.toString().toLowerCase().includes("venue"));
    if (venueIndex === -1) return null;

    for (let i = 1; i < data.length; i++) {
      const row = data[i].join(' ').toLowerCase();
      if (row.includes(name.toLowerCase()) || row.includes(reg.toLowerCase())) {
        Logger.log(data[i][venueIndex]);
        return data[i][venueIndex];
      }
    }
    return null;

  } catch (error) {
    Logger.log("Error in extractfromlist: " + error.toString());
    return null;
  } finally {
    // Clean up the converted file
    if (tempFile) {
      try {
        DriveApp.getFileById(tempFile.id).setTrashed(true);
      } catch (e) {
        Logger.log("Error deleting temp file: " + e.toString());
      }
    }
  }
}
