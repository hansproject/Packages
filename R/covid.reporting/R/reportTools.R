#' @export
adjustLength= function(aList, lists){

  tmp_list <- aList
  all_lengths <- c()

  for(elt in lists){
    all_lengths <- c(all_lengths, length(elt))
  }

  for(i in length(tmp_list):
      max(all_lengths)){
    if(i < max(all_lengths)){
      tmp_list  <- c(tmp_list,"")
    }
  }
  return(tmp_list)

}

#' @export
connectServer = function(chrome_driver, port){
  rD <- RSelenium::rsDriver(chromever = chrome_driver, port= as.integer(port))
  Sys.sleep(1.5)
  remDr <- rD[["client"]]
  return(remDr)
}

#' @export
createReport= function(destination_path, no_temp_names, no_quest_names, no_cont_names, tests_sheet){
  max_coln = 10
  sheet_names <- c("Daily_Requirements","Overdue")
  report_sheets <- list(
    "Daily_Requirements"= list("No_Temperature"= no_temp_names,
                               "No_Questionnaire"=no_quest_names,
                               "No_Contact_Tracing"=no_cont_names),
    "Overdue"= tests_sheet)

  wb <- xlsx::createWorkbook(type="xlsx")
  # Create a sheet in that workbook to contain the data table
  for(name in sheet_names){
    sheet <- xlsx::createSheet(wb, sheetName = name)
    # TABLE_ROWNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE)
    TABLE_COLNAMES_STYLE <- xlsx::CellStyle(wb) + xlsx::Font(wb, isBold=TRUE) +
      xlsx::Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER") +
      xlsx::Border(color="black", position=c("TOP", "BOTTOM"),
             pen=c("BORDER_THIN", "BORDER_THICK"))
    xlsx::setColumnWidth(sheet, colIndex=c(1:max_coln), colWidth=max(nchar(names(report_sheets)))*1.25)
    cur_report <- report_sheets[name]
    names(cur_report)<- NULL
    # Add header
    xlsx::addDataFrame(cur_report, sheet, startRow=2, startColumn=1, colnamesStyle = TABLE_COLNAMES_STYLE, row.names = FALSE)
  }
  xlsx::saveWorkbook(wb, paste(destination_path,month.name[lubridate::month(Sys.Date())],"/Daily Report ",format(Sys.Date(),"%m-%d-%Y"),".xlsx",sep = ""))
}

#' @export
downloadCovidReport= function(url, remDr){
  remDr$navigate(url)
  # username
  user_name <- remDr$findElement(using = "id", value = "username")
  user_name$clearElement()
  user_name$sendKeysToElement(sendKeys = list("hndeffo@ibrainnyc.org"))

  # password
  pwd <- remDr$findElement(using = "id", value = "password")
  pwd$clearElement()
  pwd$sendKeysToElement(sendKeys = list("HP.Lionel1"))
  sign_in <- remDr$findElement(using = "id", value = "submit_btn")
  sign_in$clickElement()
  Sys.sleep(4.25)
  cvd_dashboard <- remDr$findElement(using = "link text", value = "Covid Dashboard")
  cvd_dashboard$clickElement()
  downloaded_file <- paste("C:/Users/ndeff/Downloads/covid-details-",Sys.Date(),".xlsx",sep = "")
  if (file.exists(downloaded_file)) {
    #Delete file if it exists
    file.remove(downloaded_file)
  }

  # Check If the folder for the report exists
  output_dir <- file.path("C:/Users/ndeff/OneDrive/Desktop/Research and online classes/Company Projects/iBrain/Programs/",
                          month.name[lubridate::month(Sys.Date())])

  if (!dir.exists(output_dir)){
    dir.create(output_dir)
  } else {
    print("Dir already exists!")
  }

  Sys.sleep(4)
  cvd_dashboard <- remDr$findElement(using = "link text", value = "Daily Spreadsheet")
  cvd_dashboard$clickElement()

}

#' @export
getNoTempReport= function(file_path){
  tmp_log <- readxl::read_excel(file_path, sheet = "Temperature Log")

  dates      <- c()
  emp_names  <- c()
  emp_emails <- c()

  for(i in 1:dim(tmp_log)[1]){
    if(!tmp_log$Employee[i] %in% getExemptList())
      if(is.na(tmp_log$Temperature[i])){
        dates <- c(dates, tmp_log$Date[i])
        emp_names <- c(emp_names, tmp_log$Employee[i])
        emp_emails <- c(emp_emails, tmp_log$Email[i])
      }

  }

  # To create empty sheet in case nobody is found
  dates <- c(dates,"")
  emp_names<- c(emp_names, "")
  emp_emails <- c(emp_emails, "")

  no_tempData_sheet = list(
    "Date"= dates,
    "Names"= emp_names,
    "Emails"= emp_emails
  )

  return(no_tempData_sheet)
}

#' @export
getNoQuesReport= function(file_path){
  tmp_log <- readxl::read_excel(file_path, sheet = "Temperature Log")
  quest_log <- readxl::read_excel(file_path, sheet = "Questionnaire Detail")

  dates      <- c()
  emp_names  <- c()
  emp_emails <- c()

  for(i in 1:dim(quest_log)[1]){
    if(!quest_log$Employee[i]%in% getExemptList())
      if(is.na(quest_log$`Has Symptoms`[i])
         ||is.na(quest_log$`Contact with someone with symptoms`[i])
         ||is.na(quest_log$`Contact with someone with COVID`[i])
         ||is.na(quest_log$`Traveled to a restricted area`[i])){
        dates <- c(dates, quest_log$Date[i])
        emp_names <- c(emp_names, quest_log$Employee[i])
        emp_emails <- c(emp_emails, tmp_log$Email[match(quest_log$Employee[i],tmp_log$Employee)])
      }

  }

  # To create empty sheet in case nobody is found
  dates <- c(dates,"")
  emp_names<- c(emp_names, "")
  emp_emails <- c(emp_emails, "")

  no_quest_sheet = list(
    "Date"= dates,
    "Names"= emp_names,
    "Emails"= emp_emails
  )

  return(no_quest_sheet)
}

#' @export
getNoContactReport= function(file_path){
  report_log <- readxl::read_excel(file_path, sheet = "Last Report Dates")

  dates      <- c()
  emp_names  <- c()
  emp_emails <- c()

  for(i in 1:dim(report_log)[1]){
    if(!report_log$Employee[i]%in% getExemptList()){
      if(!is.na(report_log$`Last Contact Trace Date`[i])){
        if(!Sys.Date() == as.Date(report_log$`Last Contact Trace Date`[i],format = "%m/%d/%Y")&& !(Sys.Date()-1) == as.Date(report_log$`Last Contact Trace Date`[i],format = "%m/%d/%Y")){
          dates <- c(dates, report_log$`Last Contact Trace Date`[i])
          emp_names <- c(emp_names, report_log$Employee[i])
          emp_emails <- c(emp_emails, report_log$Email[match(report_log$Employee[i],report_log$Employee)])
        }
      }
      if(is.na(report_log$`Last Contact Trace Date`[i])){
        dates <- c(dates, "")
        emp_names <- c(emp_names, report_log$Employee[i])
        emp_emails <- c(emp_emails, report_log$Email[match(report_log$Employee[i],report_log$Employee)])
      }
    }
  }

  # To create empty sheet in case nobody is found
  dates <- c(dates,"")
  emp_names<- c(emp_names, "")
  emp_emails <- c(emp_emails, "")

  cont_sheet = list(
    "Date"= dates,
    "Names"= emp_names,
    "Emails"= emp_emails
  )
  return(cont_sheet)
}

#' @export
getOverdueReport= function(file_path){
  tests_log <- readxl::read_excel(file_path, sheet = "COVID Tests")

  dates      <- c()
  no_test_emp_names  <- c()
  no_pcr_emp_names <- c()
  pending_no_result <- c()
  overdue_emp_names <- c()
  days_pending <- c()
  days_overdue       <- c()
  emp_emails <- c()

  for(i in 1:dim(tests_log)[1]){
    if(!tests_log$Employee[i]%in% getExemptList()){

      if(!is.na(tests_log$`Days Since Last Test`[i])){
        if(tests_log$`Days Since Last Test`[++i]>= 15){
          #dates <- c(dates, tests_log$`Last Test Date`[i])
          days_overdue <- c(days_overdue, tests_log$`Days Since Last Test`[++i]-14)
          overdue_emp_names <- c(overdue_emp_names, tests_log$Employee[i])
          emp_emails <- c(emp_emails, tests_log$Email[match(tests_log$Employee[i],tests_log$Employee)])
        }
      }

      if(is.na(tests_log$`Days Since Last Test`[i])){
        #dates <- c(dates,"No Test Uploaded")
        no_test_emp_names <- c(no_test_emp_names, tests_log$Employee[i])
        emp_emails <- c(emp_emails, tests_log$Email[match(tests_log$Employee[i],tests_log$Employee)])
      }

      if(!is.na(tests_log$`Last Test Date`[i])){
        if(is.na(tests_log$`Last Result Date`[i])&& pracma::strcmp(tests_log$`Last Result`[i], "Pending")){
          pending_no_result <- c(pending_no_result, tests_log$Employee[i])
          emp_emails <- c(emp_emails, tests_log$Email[match(tests_log$Employee[i],tests_log$Employee)])
          days_pending <- c(days_pending, as.numeric(Sys.Date()-as.Date(tests_log$`Last Test Date`[i], format = "%m/%d/%Y")))
        }
      }
      if(!is.na(tests_log$`Last Test Type`[i])){
        if(pracma::strcmp(tests_log$`Last Test Type`[i], "Rapid")){
          no_pcr_emp_names <- c(no_pcr_emp_names, tests_log$Employee[i])
          emp_emails <- c(emp_emails, tests_log$Email[match(tests_log$Employee[i],tests_log$Employee)])
        }
      }
    }
  }


  for(i in length(no_test_emp_names):
      max(length(no_test_emp_names),length(no_pcr_emp_names),length(pending_no_result),length(overdue_emp_names),length(days_pending))){
    if(i <max(length(no_test_emp_names),length(no_pcr_emp_names),length(pending_no_result),length(overdue_emp_names),length(days_pending)))   {
      no_test_emp_names  <- c(no_test_emp_names,"")
    }
  }

  for(i in length(no_pcr_emp_names):
      max(length(no_test_emp_names),length(no_pcr_emp_names),length(pending_no_result),length(overdue_emp_names),length(days_pending))){
    if(i <max(length(no_test_emp_names),length(no_pcr_emp_names),length(pending_no_result),length(overdue_emp_names),length(days_pending))){
      no_pcr_emp_names  <- c(no_pcr_emp_names,"")
    }
  }

  for(i in length(pending_no_result):
      max(length(no_test_emp_names),length(no_pcr_emp_names),length(pending_no_result),length(overdue_emp_names),length(days_pending))) {
    if(i <max(length(no_test_emp_names),length(no_pcr_emp_names),length(pending_no_result),length(overdue_emp_names),length(days_pending))){
      pending_no_result  <- c(pending_no_result,"")
      days_pending  <- c(days_pending, "")
    }
  }

  for(i in length(overdue_emp_names):
      max(length(no_test_emp_names),length(no_pcr_emp_names),length(pending_no_result),length(overdue_emp_names),length(days_pending))){
    if(i <max(length(no_test_emp_names),length(no_pcr_emp_names),length(pending_no_result),length(overdue_emp_names),length(days_pending))){
      overdue_emp_names  <- c(overdue_emp_names,"")
      days_overdue       <- c(days_overdue, "")
    }
  }

  tests_sheet = list(
    "No Test"= no_test_emp_names,
    "No PCR Uploaded"= no_pcr_emp_names,
    "Pending No Result"=pending_no_result,
    "Days Pending"= days_pending,
    "Names"= overdue_emp_names,
    "Days Overdue"=days_overdue

  )

  return(tests_sheet)
}

#' @export
getExemptList=function(){
  return(c("Victor Pedro","Patrick Donohue","Linda Brower","Katandria Johnson",
    "John Cardenas","Michael Geraci","Alisa Rockett",
    "Bethany Barber","Faith Okkunrobo","Frieda Doohkey","Jeff Jean-Louis","Suzanne Wallach",
    "Bridget Halverson","Lise Joseph","Marisa Motley","Annette Johnson", "Ron Rivas", "Viana Jean-Pierre", "Deena Coorie","Niurca Salazar"))
}

#' @export
getManagementEmails= function(){
  return(c("katandria@ibrainnyc.org","swallach@ibrainnyc.org","akristofik@ibrainnyc.org", "john@ibrainnyc.org"))
}

#' @export
getDailyReqEmail=function(){
  cur_time <- ""

  if(lubridate::hour(Sys.time())< 18){
    cur_time <- "10:00am"
    body    <- "Good morning,
I hope you are doing well. I attached to this email the Covid Overdue Report + 3 Requirements Daily Report.
Best regards
"
  }else{
    cur_time <- "6:00pm"
    body    <- "Good evening,
I hope you are doing well. I attached to this email the Covid Overdue Report + 3 Requirements Daily Report.
Best regards
"
  }
  subject <- paste(cur_time," Daily Report")
  return(list("body"= body, "subject"= subject))
}

#' @export
sendEmail= function (from, to, subject, body, bcc, cc, file_path){
  from <- from
  to <- to
  subject <- subject
  body <- body

  mailR::send.mail(from=from,to=to,subject=subject,body=body, bcc= bcc, cc= cc,smtp = list(host.name = "smtp.office365.com", port = 587,
                                                                                    user.name="hndeffo@ibrainnyc.org", passwd="HPcameroun1",tls= TRUE,ssl=FALSE),
            authenticate = TRUE, attach.files= c(paste(file_path,month.name[lubridate::month(Sys.Date())],"/Daily Report ",format(Sys.Date(),"%m-%d-%Y"),".xlsx",sep = "")), send = TRUE)
}
