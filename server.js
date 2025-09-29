const my_name="Your_Name";
const my_reg_number="Your_Registration_Number"; //22bce****


function myFunction() {
  const now=new Date();
  Logger.log(now);
  Logger.log("Checking in last 30 minutes");
  const thirtyMinutesAgo=new Date(now.getTime()-30*60*1000);
  Logger.log(thirtyMinutesAgo);

 
   const query="(from:"placement_cell_email"  OR from:"placement_cell_email" OR to:"placement_cell_email") is:unread"; // from: abc@gmail.com


  const threads=GmailApp.search(query);
  const recentThreads= threads.filter(thread=> thread.getLastMessageDate()>= thirtyMinutesAgo);

  if(recentThreads.length===0) return;
  Logger.log(recentThreads);

const createdEvents = new Set();
const processedMessages = new Set();

  for(let thread of recentThreads) {
   
    const messages=thread.getMessages();
    for(let msg of messages){
      
              if (processedMessages.has(msg.getId())) continue; // skip if already processed
        processedMessages.add(msg.getId());               // mark as processed

   
      const subject=msg.getSubject();

      // if(!(subject.toLowerCase().includes("talk") ||
      //    subject.toLowerCase().includes("test") ||
      //    subject.toLowerCase().includes("process") ||
      //    subject.toLowerCase().includes("online")) ){
      //    Logger.log("No test ");
      //    continue;}
      const keywords = ["talk", "test", "process", "online", "assessment", "exam"];
        if (!keywords.some(k => subject.toLowerCase().includes(k))) {
            Logger.log("No test");
            continue;
        }

      const attachments=msg.getAttachments();
      for(let attachment of attachments){
        if(attachment.getContentType().includes("sheet")|| attachment.getName().match(/\.(xlsx|xls)$/)){
          const blob=attachment.copyBlob();
         

      const fileMetadata = {
            'name': attachment.getName().replace(/\.(xlsx|xls)$/, ''),
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        };
        
        // Use the blob directly, not getBytes()
        const file = Drive.Files.create(fileMetadata, blob);
           Logger.log("Adding candidates list into drive")
        
        const sheetfile = SpreadsheetApp.openById(file.id);

          let found=false;
      

                  for (let sheet of sheetfile.getSheets()) {
          const data = sheet.getDataRange().getValues();
          for (let row of data) {
            if (row.join(' ').toLowerCase().includes(my_name.toLowerCase()) || row.join(' ').toLowerCase().includes(my_reg_number.toLowerCase())) {
              found = true;
              break; // exit row loop
            }
          }
          if (found) break; // exit sheet loop
        }


           Logger.log("Deleting candidates list from drive")
          DriveApp.getFileById(file.id).setTrashed(true);

          if(found){

            const basicEventKey = `${subject}`;
  
  if (createdEvents.has(basicEventKey)) {
    Logger.log("Event already created, skipping duplicate - no extraction needed");
    continue;
  }
            let venue="";
            if(msg.getSubject().toLowerCase().includes("own location")){
              venue="Own Location";
            }

            else venue=extractvenue(attachments)||"TBD";
          
            let start=extractDateTimeFromText(msg.getSubject()+msg.getPlainBody());


            if (!start) {
              Logger.log("No date found, using tomorrow 10:00 AM as default");
              const tomorrow = new Date();
              tomorrow.setDate(tomorrow.getDate() + 1);
              tomorrow.setHours(10, 0, 0, 0); // 10:00 AM
              start = tomorrow;
            }
            Logger.log(start);

            //Default 1hr
            const end=new Date(start.getTime()+60*60*1000);


            const eventKey = `${msg.getSubject()}-${start.getTime()}`;
            if (createdEvents.has(eventKey)) {
                Logger.log("Event already created, skipping duplicate");
                continue;
              }

            //Creating Calendar Event
                        try{ CalendarApp.createEvent(
                        msg.getSubject(),
                        start,
                        end,

                        {
              location: `${venue}`,
              description: `
            📍 Venue: ${venue}
            ⏱️ Duration: ${Math.round((end.getTime() - start.getTime()) / (1000 * 60))} minutes

            📧 Original Email: ${msg.getThread().getPermalink()}

            💡 Remember: Bring ID card & arrive 15 mins early!


              `,
              
              // Just essential reminders
              reminders: {
                useDefault: false,
                overrides: [
                  {method: 'popup', minutes: 60},   // 1 hour before
                  {method: 'popup', minutes: 15}    // 15 minutes before  
                ]
              }
            }
            
            
          );
          createdEvents.add(eventKey);
          
              if(venue==="TBD") Logger.log("Created calendar event with TBD");
                else Logger.log("Created calendar event with location");
          }catch(error){
              Logger.log("Error creating calendar event " + error.toString());
          }
          }
        }
      }
    }
  }
  // Clear set for memory efficiency
  createdEvents.clear();
  processedMessages.clear();
  Logger.log("Cleared event tracking set for memory efficiency");

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
  /on\s+(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/i ,                 // on 08/09/2025 or on 08-09-2025
];

    for(let i=0;i<patterns.length;i++){
      let match= text.match(patterns[i]);
      if(match){
        let date,month,year,formattedDate;

        if(i==0 || i==3){
          date=match[1];
          month=match[2];
          year=match[3];
          formattedDate=`${month} ${date} ${year}`;
        }
        else if(i==1){
           date=match[1];
          month=match[2];
          year=match[3];
          formattedDate=`${month} ${date} ${year}`;
        }
        else if(i==2){
          date=match[1];
          const short=match[2];
          year=match[3];

          const fullmonth=monthMap[short.toLowerCase()];
          if(!fullmonth) continue;
          formattedDate=`${fullmonth} ${date} ${year}`;
        }
        Logger.log("Extracted Date:"+formattedDate);
       // const finalDate=new Date(formattedDate);
       /* if(!isNaN(finalDate.getTime())){
          return formattedDate;
        }*/
        return formattedDate;

        
      }

    }


      return null;
        }catch(error){
          Logger.log("Error extracting date");
          return null;
        }  
}



function extractTime(text){
  // try{
  //   Logger.log(`Extracting time from :${text}`);

  //  // const pattern1= /(\d{1,2})[:.](\d{2})\s*([ap])\.?m\.?/i;
  //   const pattern1 = /(\d{1,2})[:.](\d{2})\s*(am|pm)/i;
  //   const timePattern = /(?:by\s+)?(\d{1,2})[:.](\d{2})\s*([ap])\.?m\.?/i;



  //   let match=text.match(pattern1);

  //   if(match){
  //     const hour=match[1];
  //     const minute=match[2];
  //     const ampm=match[3].toUpperCase();
  //     const time=`${hour}:${minute} ${ampm}`;
  //     Logger.log("found pattern1 " +time);

  //     return time;
  //   }
  // const pattern2 = /by\s+(\d{1,2})[:.](\d{2})\s*(am|pm)/i;
  //   match = text.match(pattern2);

  //   if (match) {
  //     const hour = match[1];
  //     const minute = match[2];
  //     const ampm = match[3].toUpperCase();
  //     const time = `${hour}:${minute} ${ampm}`;
  //     Logger.log(`Found time pattern 2: ${time}`);
  //     return time;
  //   }

  //   return null;

  // }catch(error){
  //   Logger.log("error extracting time:" +error.toString());
  //   return null;
  // }
   try {
    Logger.log(`Extracting time from: ${text}`);

    // Single robust pattern for almost all common styles
   // const timePattern = /(?:by\s+)?(\d{1,2})[:.](\d{2})\s*([ap])\.?m\.?/i;
    const timePattern = /(?:by\s+)?(\d{1,2})[:.](\d{2})\s*([ap]m)/i;


    const match = text.match(timePattern);

    if (match) {
      const hour = match[1];
      const minute = match[2];
      const ampm = match[3].toUpperCase();
      const time = `${hour}:${minute} ${ampm}`;
      Logger.log("Found time: " + time);
      return time;
    }

    return null;
  } catch (error) {
    Logger.log("Error extracting time: " + error.toString());
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
          Logger.log("found venue in list(venues list)");
          return extractedVenue;
        }
      }
      Logger.log("checking other attachments")
      extractedVenue=extractfromlist(attachment,my_name,my_reg_number);
        if(extractedVenue){
          Logger.log("found venue in list");
          return extractedVenue;
        }

    }
    return null;
  }catch(error){
    Logger.log("Error in venue");
    return null;
  }
}





function extractfromlist(attachment, name, reg) {
  let tempFile = null;
  try {
    const blob = attachment.copyBlob();

    // Convert uploaded Excel file to Google Sheet
    const fileResource = {
      title: attachment.getName(),
      mimeType: MimeType.GOOGLE_SHEETS
    };
    tempFile = Drive.Files.create(fileResource, blob, {convert: true});

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
        Logger.log("deleting venues list from drive");

        DriveApp.getFileById(tempFile.id).setTrashed(true);
      } catch (e) {
        Logger.log("Error deleting temp file: " + e.toString());
      }
    }
  }
}
