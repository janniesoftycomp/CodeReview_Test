Main                 PROCEDURE                             ! Declare Procedure

                    MAP
GetConfig              FUNCTION(STRING,STRING),STRING
PutConfig              PROCEDURE(STRING,STRING)
                    END

GLO:StagePackage1Description string(128)
GLO:StagePackage2Description STRING(128)
GLO:StagePackage3Description STRING(128)
GLO:StagePackage4Description STRING(128)

GLO:DashboardQueryType      LONG(0)
GLO:DashboardThread LONG(0)
GLO:MessageCredit Long
ExecuteDashboard            Long

FromDate                    Long
ToDate Long

CeremonyQueue        QUEUE,PRE(CQ)                         ! 
CeremonyNo           LONG                                  ! 
                     END                                   ! 
SizeQtyQueue         QUEUE,PRE(SQQ)                        ! 
SizeDescription      STRING(20)                            ! 
Qty                  REAL                                  ! 
                     END                                   ! 


LOC:QtyGrads         LONG                                  ! 
LOC:StageOrdered     LONG                                  ! 
LOC:FamilyOrdered    LONG                                  ! 
LOC:FamilyQtyA       LONG                                  ! 
LOC:FamilyQtyB       LONG                                  ! 
LOC:CashTotal        DECIMAL(14,2)                         ! 
LOC:CreditTotalTotal DECIMAL(14,2)                         ! 
LOC:ChequeTotal      DECIMAL(14,2)                         ! 
LOC:ElectronicTotal  DECIMAL(14,2)                         ! 
LOC:GrandTotal       DECIMAL(14,2)                         ! 
LOC:RationPerGrad    DECIMAL(14,2)                         ! 
LOC:CountedPhotostaken LONG                                ! 

LOC:LastUpdateString STRING(80)                            ! 
LOC:CountofBankBoards LONG                                 ! 
LOC:PrintServiceCount LONG
LOC:SMSCount         LONG                                  ! 
LOC:EmailCount       LONG                                  ! 
LOC:NoOfOrdersClients LONG                                 ! 
LOC:NoOfNew          LONG                                  ! 
LOC:NoOfInProcess    LONG                                  ! 
LOC:NoOfPosted       LONG                                  ! 
LOC:NoOfCancelled    LONG                                  ! 
LOC:NoOfReturned     LONG                                  ! 
LOC:NoOfCollected    LONG                                  ! 
LOC:NoOfCdPackages   LONG                                  ! 
LOC:NoOfJobcards     LONG                                  ! 
LOC:NoOfJobcardsOpen LONG                                  ! 
LOC:NoOfJobcardsLab  LONG                                  ! 
LOC:NoOfJobcardsCompleted LONG                             ! 
LOC:NoOfJobcardsCollected LONG                             ! 
LOC:NoOfJobcardsPosted LONG                                ! 
LOC:NoOfJobcardsReturned LONG                              ! 
LOC:QueryType        STRING(40)                            ! 
LOC:StageOrderLine1Total DECIMAL(10,2)                     ! 
LOC:StageOrderLine2Total DECIMAL(10,2)                     ! 
LOC:StageOrderLine3Total DECIMAL(10,2)                     ! 
LOC:StageOrderLine4Total DECIMAL(10,2)                     ! 
LOC:HighresCDQty     LONG                                  ! 
LOC:HighresCDValue   DECIMAL(7,2)                          ! 
LOC:FamilyCalcTotal  DECIMAL(10,2)                         ! 
LOC:HighResFamilyCDOrdered LONG                            ! 
LOC:StageLine1Qty    LONG                                  ! 
LOC:StageLine2Qty    LONG                                  ! 
LOC:StageLine3Qty    LONG                                  ! 
LOC:PostalCalcTotal  DECIMAL(10,2)                         ! 
LOC:StageLine4Qty    STRING(20)                            ! 
LOC:SchoolTotal      DECIMAL(14,2)                         ! 
LOC:PreSchoolTotal   DECIMAL(14,2)                         ! 
LOC:SumOrderedTotals DECIMAL(14,2)                         ! 
LOC:DocumentFromDate DATE                                  ! 
LOC:DocumentToDate   DATE                                  ! 
LOC:CountFullShorts  LONG                                  ! 
LOC:CountInOrderShort LONG                                 ! 
LOC:InstituteNo      LONG                                  ! 
LOC:InstituteDescription STRING(80)                        ! 
LOC:CurrentAction    STRING(255)                           ! 
LOC:Type             LONG                                  ! 
LOC:ReturnedMail2008 LONG                                  ! 
LOC:ReturnedMail2009 LONG                                  ! 
LOC:ReturnedMail2010 LONG                                  ! 
LOC:ReturnedMail2011 LONG                                  ! 
LOC:ReturnedMail2012 LONG                                  ! 
LOC:StageQty         LONG                                  ! 
LOC:FamilyA3Qty      LONG                                  ! 
LOC:FamiliyA4Qty     LONG                                  ! 
LOC:StageA4Qty       LONG                                  ! 
LOC:StageJumboQty    LONG                                  ! 



LocProc         Class
Init                                Procedure()
ExecuteFromAndToDateQuery           Procedure()
ExecuteDateandInstituteQuery        Procedure()
ExecuteInstituteQuery               Procedure()
ExecuteAllCombinedQuery             Procedure()
RefreshScreens                      Procedure()
ClearAll                            Procedure()
CalculateQtyViaSize                 Procedure()
ExecuteAllCombinedQuery2016         PROCEDURE()
ExecuteFromAndToDateQuery2016       Procedure()
ExecuteDateandInstituteQuery2016    Procedure()
ExecuteInstituteQuery2016           Procedure()
ClearAll2016                        Procedure()
ExecuteJobcardQuerySchool           Procedure()
ExecuteJobcardQueryPreSchool        Procedure()
ExecuteDocumentQuery        		Procedure()
ExecuteShortsQuery                	Procedure()
LoadPhotosTakenDateInstitute        PROCEDURE()
LoadPhotosTakenInstitute 			PROCEDURE()
LoadPhotosTakenDate 				PROCEDURE()
                END

Window                      WINDOW('Caption'),AT(,,395,224),GRAY,FONT('Segoe UI',9)
                                BUTTON('&OK'),AT(1,1,10,10),USE(?Button5:3)
                                BUTTON('&OK'),AT(1,1,10,10),USE(?Button6)
                                BUTTON('&OK'),AT(1,1,10,10),USE(?Button5)
                                BUTTON('&OK'),AT(1,1,10,10),USE(?Button6:2)
                                BUTTON('&OK'),AT(1,1,10,10),USE(?Button5:2)
                                BUTTON('&OK'),AT(1,1,10,10),USE(?Button5:4)
                                BUTTON('&OK'),AT(1,1,10,10),USE(?Button6:3)
                                BUTTON('&OK'),AT(1,1,10,10),USE(?Button3)
                                BUTTON('&OK'),AT(1,1,10,10),USE(?Button13)
                                PROMPT('X'),AT(20,1,10,10),USE(?LOC:StageOrderLine1Total:Prompt)
                                PROMPT('X'),AT(20,1,10,10),USE(?LOC:StageOrderLine2Total:Prompt)
                                PROMPT('X'),AT(20,1,10,10),USE(?LOC:StageOrderLine3Total:Prompt)
                                PROMPT('X'),AT(20,1,10,10),USE(?LOC:StageOrderLine4Total:Prompt)
                            END
SQLScript            FILE,DRIVER('MSSQL'),PRE(SQL),BINDABLE,CREATE,THREAD,EXTERNAL(''),DLL(dll_mode) !                     
Record                   RECORD,PRE()
Field1                      CSTRING(256)                   !                     
Field2                      CSTRING(256)                   !                     
Field3                      CSTRING(256)                   !                     
Field4                      CSTRING(256)                   !                     
Field5                      CSTRING(256)                   !                     
Field6                      CSTRING(256)                   !                     
Field7                      CSTRING(256)                   !                     
Field8                      CSTRING(256)                   !                     
Field9                      CSTRING(256)                   !                     
Field10                     CSTRING(256)                   !                     
Field11                     CSTRING(256)                   !                     
Field12                     CSTRING(256)                   !                     
Field13                     CSTRING(256)                   !                     
Field14                     CSTRING(256)                   !                     
Field15                     CSTRING(256)                   !                     
Field16                     CSTRING(256)                   !                     
Field17                     CSTRING(256)                   !                     
Field18                     CSTRING(256)                   !                     
Field19                     CSTRING(256)                   !                     
Field20                     CSTRING(256)                   !                     
Field21                     CSTRING(256)                   !                     
Field22                     CSTRING(256)                   !                     
Field23                     CSTRING(256)                   !                     
Field24                     CSTRING(256)                   !                     
Field25                     CSTRING(256)                   !                     
Field26                     CSTRING(256)                   !                     
Field27                     CSTRING(256)                   !                     
Field28                     CSTRING(256)                   !                     
Field29                     CSTRING(256)                   !                     
Field30                     CSTRING(256)                   !                     
                            END
                        END 

SQLScript2           FILE,DRIVER('MSSQL'),PRE(SQLS),BINDABLE,CREATE,THREAD,EXTERNAL(''),DLL(dll_mode) !                     
Record                   RECORD,PRE()
Field1                      CSTRING(256)                   !                     
Field2                      CSTRING(256)                   !                     
Field3                      CSTRING(256)                   !                     
Field4                      CSTRING(256)                   !                     
Field5                      CSTRING(256)                   !                     
Field6                      CSTRING(256)                   !                     
Field7                      CSTRING(256)                   !                     
Field8                      CSTRING(256)                   !                     
Field9                      CSTRING(256)                   !                     
Field10                     CSTRING(256)                   !                     
                         END
                     END                       
SQLScript3           FILE,DRIVER('MSSQL'),PRE(SQLSS),BINDABLE,CREATE,THREAD,EXTERNAL(''),DLL(dll_mode) !                     
Record                   RECORD,PRE()
Field1                      CSTRING(256)                   !                     
Field2                      CSTRING(256)                   !                     
Field3                      CSTRING(256)                   !                     
Field4                      CSTRING(256)                   !                     
Field5                      CSTRING(256)                   !                     
Field6                      CSTRING(256)                   !                     
Field7                      CSTRING(256)                   !                     
Field8                      CSTRING(256)                   !                     
Field9                      CSTRING(256)                   !                     
Field10                     CSTRING(256)                   !                     
                         END
                     END                       
SQLScript8           FILE,DRIVER('MSSQL'),PRE(SQLS8),BINDABLE,CREATE,THREAD,EXTERNAL(''),DLL(dll_mode) !                     
Record                   RECORD,PRE()
Field1                      CSTRING(256)                   !                     
Field2                      CSTRING(256)                   !                     
Field3                      CSTRING(256)                   !                     
Field4                      CSTRING(256)                   !                     
Field5                      CSTRING(256)                   !                     
                         END
                     END                       



Dashboard2016Alias   FILE,DRIVER('TOPSPEED'),PRE(Dasha),CREATE,BINDABLE,THREAD,EXTERNAL('')!,DLL(dll_mode) !                     
Dashboard2016_PK         KEY(Dasha:Dashboard2016No),NOCASE,PRIMARY !                     
Record                   RECORD,PRE()
Dashboard2016No             LONG                           !                     
PackageNo                   LONG                           !                     
PackageDescription          STRING(1210)                   !                     
TotalQty                    DECIMAL(14,2)                  !                     
TotalValue                  DECIMAL(14,2)                  !                     
                         END
                        END 
Dashboard2016        FILE,DRIVER('TOPSPEED'),PRE(Dash),CREATE,BINDABLE,THREAD,EXTERNAL('')!,DLL(dll_mode) !                     
Dashboard2016_PK         KEY(Dash:Dashboard2016No),NOCASE,PRIMARY !                     
Record                   RECORD,PRE()
Dashboard2016No             LONG                           !                     
PackageNo                   LONG                           !                     
PackageDescription          STRING(1210)                   !                     
TotalQty                    DECIMAL(14,2)                  !                     
TotalValue                  DECIMAL(14,2)                  !                     
                         END
                     END                       


LPhotoSize           FILE,DRIVER('MSSQL'),PRE(LPHOSIZ),BINDABLE,CREATE,THREAD,EXTERNAL('')!,DLL(dll_mode) !                     
PhotoSize_PK             KEY(LPHOSIZ:SysIdPhotoSize),NAME('LPHOSIZ_PK_PhotoSize'),PRIMARY !                     
PhotoSize_SK_Description KEY(LPHOSIZ:DescriptionPhotoSize),DUP,NAME('SK_PhotoSize_Description'),NOCASE,OPT !                     
Record                   RECORD,PRE()
SysIdPhotoSize              LONG                           ! PhotoSize           
DescriptionPhotoSize        CSTRING(51)                    ! Description         
WidthtInmm                  LONG                           !                     
LengthInmm                  LONG                           !                     
AllowForWebOrder            LONG                           !                     
                         END
                     END                      

Institute            FILE,DRIVER('MSSQL'),PRE(Ins),BINDABLE,CREATE,THREAD,EXTERNAL('')!,DLL(dll_mode) ! Institute           
PK_Institute             KEY(Ins:SysIDInstitute),NAME('PK_Institute'),PRIMARY ! PK_Institute        
SK_Institute_Name        KEY(Ins:NameInstitute),DUP,NAME('SK_Institute_Name') ! SK_Institute_Name   
FK_Institute_TaxType     KEY(Ins:SysIdTaxType),DUP,NAME('FK_Institute_TaxType'),NOCASE,OPT ! FK_Institute_TaxType
FK_Institute_Tax         KEY(Ins:SysIDTax),DUP,NAME('FK_Institute_Tax') ! FK_Institute_Tax    
Institute_FK_BOA         KEY(Ins:BoardNo),DUP,NOCASE,OPT   !                     
Institute_FK_PCK         KEY(Ins:DefaultPackageNo),DUP,NOCASE,OPT !                     
Record                   RECORD,PRE()
SysIDInstitute              LONG                           ! System ID Institute 
ShortNameInstitute          CSTRING(21)                    !                     
NameInstitute               CSTRING(51)                    ! Name                
SysIdDepartment             LONG                           ! DepartmentID        
SysIDTax                    LONG                           ! System ID Tax       
SysIdTaxType                LONG                           ! TaxTypeID - Inclusive/Exclusive/Exempt
UserCreate                  CSTRING(13)                    ! UserCreate          
BoardNo                     LONG                           !                     
DateCreate                  STRING(8)                      ! DateCreate          
DateCreate_GROUP            GROUP,OVER(DateCreate)         ! DateCreate_GROUP    
DateCreate_DATE               DATE                         ! DateCreate          
TimeCreate_DATE               TIME                         ! TimeCreate          
                            END                            !                     
UserChange                  CSTRING(13)                    ! UserChange          
DateChange                  STRING(8)                      ! DateChange          
DateChange_GROUP            GROUP,OVER(DateChange)         ! DateChange_GROUP    
DateChange_DATE               DATE                         ! Date_DATE           
TimeChange_DATE               TIME                         ! TimeChange          
                            END                            !                     
AmountPerHead               DECIMAL(7,2)                   !                     
StageLine1Cost              DECIMAL(7,2)                   !                     
StageLine2Cost              DECIMAL(7,2)                   !                     
StageLine3Cost              DECIMAL(7,2)                   !                     
StageLine4Cost              DECIMAL(7,2)                   !                     
StagePostagePrice           DECIMAL(7,2)                   !                     
FamilyPostagePrice          DECIMAL(7,2)                   !                     
FamiliyUpsizePrice          DECIMAL(7,2)                   !                     
FamiliyHighResPrice         DECIMAL(7,2)                   !                     
DefaultPackageDescription   STRING(40)                     !                     
DefaultPackageNo            LONG                           !                     
StageSuperSizePrice         DECIMAL(7,2)                   !                     
StageUpSizePrice            DECIMAL(7,2)                   !                     
StagePackage1Description    STRING(60)                     !                     
StagePackage2Description    STRING(60)                     !                     
StagePackage3Description    STRING(60)                     !                     
StagePackage4Description    STRING(60)                     !                     
InternationalInstitute      LONG                           !                     
                         END
                     END                       

InstituteAlias       FILE,DRIVER('MSSQL'),PRE(INSA),BINDABLE,CREATE,THREAD,EXTERNAL(''),DLL(dll_mode) !                     
PK_Institute             KEY(INSA:SysIDInstitute),NAME('PK_Institute'),PRIMARY ! PK_Institute        
SK_Institute_Name        KEY(INSA:NameInstitute),DUP,NAME('SK_Institute_Name') ! SK_Institute_Name   
FK_Institute_TaxType     KEY(INSA:SysIdTaxType),DUP,NAME('FK_Institute_TaxType'),NOCASE,OPT ! FK_Institute_TaxType
FK_Institute_Tax         KEY(INSA:SysIDTax),DUP,NAME('FK_Institute_Tax') ! FK_Institute_Tax    
Institute_FK_BOA         KEY(INSA:BoardNo),DUP,NOCASE,OPT  !                     
Institute_FK_PCK         KEY(INSA:DefaultPackageNo),DUP,NOCASE,OPT !                     
Record                   RECORD,PRE()
SysIDInstitute              LONG                           ! System ID Institute 
ShortNameInstitute          CSTRING(21)                    !                     
NameInstitute               CSTRING(51)                    ! Name                
SysIdDepartment             LONG                           ! DepartmentID        
SysIDTax                    LONG                           ! System ID Tax       
SysIdTaxType                LONG                           ! TaxTypeID - Inclusive/Exclusive/Exempt
UserCreate                  CSTRING(13)                    ! UserCreate          
BoardNo                     LONG                           !                     
DateCreate                  STRING(8)                      ! DateCreate          
DateCreate_GROUP            GROUP,OVER(DateCreate)         ! DateCreate_GROUP    
DateCreate_DATE               DATE                         ! DateCreate          
TimeCreate_DATE               TIME                         ! TimeCreate          
                            END                            !                     
UserChange                  CSTRING(13)                    ! UserChange          
DateChange                  STRING(8)                      ! DateChange          
DateChange_GROUP            GROUP,OVER(DateChange)         ! DateChange_GROUP    
DateChange_DATE               DATE                         ! Date_DATE           
TimeChange_DATE               TIME                         ! TimeChange          
                            END                            !                     
AmountPerHead               DECIMAL(7,2)                   !                     
StageLine1Cost              DECIMAL(7,2)                   !                     
StageLine2Cost              DECIMAL(7,2)                   !                     
StageLine3Cost              DECIMAL(7,2)                   !                     
StageLine4Cost              DECIMAL(7,2)                   !                     
StagePostagePrice           DECIMAL(7,2)                   !                     
FamilyPostagePrice          DECIMAL(7,2)                   !                     
FamiliyUpsizePrice          DECIMAL(7,2)                   !                     
FamiliyHighResPrice         DECIMAL(7,2)                   !                     
DefaultPackageDescription   STRING(40)                     !                     
DefaultPackageNo            LONG                           !                     
StageSuperSizePrice         DECIMAL(7,2)                   !                     
StageUpSizePrice            DECIMAL(7,2)                   !                     
StagePackage1Description    STRING(60)                     !                     
StagePackage2Description    STRING(60)                     !                     
StagePackage3Description    STRING(60)                     !                     
StagePackage4Description    STRING(60)                     !                     
InternationalInstitute      LONG                           !                     
                         END
                     END                       



  CODE



  LocProc.Init()

  Accept
      Case Accepted()
        OF ?Button5:3
              LocProc.ExecuteJobcardQuerySchool()
        OF ?Button6
              LocProc.ClearAll()
        OF ?Button5
             !ThisWindow.Update()
             ! Clear all the local variables
             LOC:QtyGrads = 0
             LOC:StageOrdered = 0
             LOC:FamilyQtyA = 0
             LOC:FamilyQtyB = 0
             LOC:FamilyOrdered = 0
             LOC:CashTotal = 0
             LOC:CreditTotalTotal = 0
             LOC:ChequeTotal = 0
             LOC:ElectronicTotal = 0
             LOC:GrandTotal = 0
             LOC:RationPerGrad = 0
          
             Case GLO:DashboardQueryType
             of 0
                 LocProc.ExecuteAllCombinedQuery()
             of 1
                 LocProc.ExecuteFromAndToDateQuery()
             of 2
                 LocProc.ExecuteInstituteQuery()
             of 3
                 LocProc.ExecuteDateandInstituteQuery()
             End

        OF ?Button6:2
              LocProc.ClearAll()

            
            
        OF ?Button5:2
             LOC:QtyGrads = 0
             LOC:StageOrdered = 0
             LOC:FamilyQtyA = 0
             LOC:FamilyQtyB = 0
             LOC:FamilyOrdered = 0
             LOC:CashTotal = 0
             LOC:CreditTotalTotal = 0
             LOC:ChequeTotal = 0
             LOC:ElectronicTotal = 0
             LOC:GrandTotal = 0
             LOC:RationPerGrad = 0
             LOC:CountedPhotostaken = 0
          
             Case GLO:DashboardQueryType
             of 0
                 LocProc.ExecuteAllCombinedQuery2016()
             of 1
                 LocProc.ExecuteFromAndToDateQuery2016()
             of 2
                 LocProc.ExecuteInstituteQuery2016()
             of 3
                 LocProc.ExecuteDateandInstituteQuery2016()
             End
     
        OF ?Button6:3
          LocProc.ClearAll2016()


        OF ?BUTTON3
              LocProc.ExecuteShortsQuery()

        OF ?Button13
              LocProc.CalculateQtyViaSize()

        END
        
        Case Event()        
          OF EVENT:Timer
              LocProc.RefreshScreens()      
          of ExecuteDashboard
             Case GLO:DashboardQueryType
             of 0
                 LocProc.ExecuteAllCombinedQuery()
             of 1
                 LocProc.ExecuteFromAndToDateQuery()
             of 2
                 LocProc.ExecuteInstituteQuery()
             of 3
                 LocProc.ExecuteDateandInstituteQuery()
             End
      
        END

    END
        
LocProc.Init    Procedure()
    CODE

    FromDate = Today()
    ToDate = Today()

    GLO:DashboardThread = Thread()

    ! Calculate the Qty of Print Jobs
    SQLScript{Prop:SQL} = 'Select Count(PrintServiceDetailNo) From PrintServiceDetail where Status = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:PrintServiceCount = SQL:Field1
    End

    ! Calculate the Qty of SMS messages to be sent
    SQLScript{Prop:SQL} = 'Select Count(CommunicationNo) From Communication where Status = 0 and OutputType = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:SMSCount = SQL:Field1
    End

    ! Calculate the Qty of Email messages to be sent
    SQLScript{Prop:SQL} = 'Select Count(CommunicationNo) From Communication where Status = 0 and OutputType = 1'
    Next(SQLScript)
    If Not Error() then
        LOC:EmailCount = SQL:Field1
    End

    ! Calculate the Qty of Print Jobs
    SQLScript{Prop:SQL} = 'Select Count(PrintServiceDetailNo) From PrintServiceDetail where Status = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:PrintServiceCount = SQL:Field1
    End

    ! Calculate the Qty of SMS messages to be sent
    SQLScript{Prop:SQL} = 'Select Count(CommunicationNo) From Communication where Status = 0 and OutputType = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:SMSCount = SQL:Field1
    End

    ! Calculate the Qty of Email messages to be sent
    SQLScript{Prop:SQL} = 'Select Count(CommunicationNo) From Communication where Status = 0 and OutputType = 1'
    Next(SQLScript)
    If Not Error() then
        LOC:EmailCount = SQL:Field1
    End

    !BRW8.ResetFromFile
    !BRW8.ResetSort(True)

    !BRW7.ResetFromFile
    !BRW7.ResetSort(True)

    ! Calculate the Overwiew Details
    SQLScript{Prop:SQL} = 'Select Count(ClientNo) From Client where Status = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfOrdersClients = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfNew = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 1'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfInProcess = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 5'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfPosted = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 6'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfCancelled = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 7'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfCollected = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 8'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfReturned = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcards = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsOpen = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 1'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsLab = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 2'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsCompleted = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 3'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsCollected = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 4'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsPosted = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 5'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsReturned = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where NedbankBoardCheck = 1'
    Next(SQLScript)
    If Not Error() then
        LOC:CountofBankBoards = SQL:Field1
    End

    LOC:LastUpdateString = 'Last Refresh date and time: ' & format(Today(),@d06b) & ' ' & Format(Clock(),@t01b)

    GLO:MessageCredit = GetConfig('GLO:MessageCredit','0')

    stream(Dashboard2016Alias)
    Dasha:Dashboard2016No = 1
    Set(Dasha:Dashboard2016_PK,Dasha:Dashboard2016_PK)
    LOOP
        !if Access:Dashboard2016Alias.Next() then break.
        !Access:Dashboard2016Alias.DeleteRecord(0)
    END

    FLUSH(Dashboard2016Alias)        
        
    Display()

LocProc.ExecuteFromAndToDateQuery       Procedure()
    CODE

    SetCursor(Cursor:Wait)

    ! The process will calculate all the combined figures

    LOC:QueryType = 'Combined Query'

    ! Calculate the Qty values
    SQLScript{Prop:SQL}= 'Select Sum(cs.NOOFGRADUATES), Sum(cs.STAGEODERQTY), Sum(cs.A4FAMILYQTY), Sum(cs.A3FAMILYQTY), Sum(cs.RATIONPERGRADUATE), sum(cs.GRANDTOTAL), Count(cs.CASHUPNO), Sum(cs.CHEQUES), Sum(cs.CREDITCARDS), Sum(cs.EFTAmount), Sum(cs.R200), Sum(cs.R100), Sum(cs.R50), Sum(cs.R20), Sum(cs.R10), Sum(cs.CHANGE), Sum(cs.StageLine1Value), Sum(cs.StageLine2Value), Sum(cs.StageLine3Value), Sum(cs.StageLine4Value), Sum(cs.TotalFamilyValue), Sum(cs.PostalValue), Sum(cs.HighResFamilyCDOrdered), Sum(cs.StageLine1Qty), sum(cs.StageLine2Qty), sum(cs.StageLine3Qty), sum(cs.StageLine4Qty) From Cashup cs where cs.cashupno in (Select distinct(cu.cashupno) From Cashup cu inner join CashupCeremony cum on (cum.CashupNo = cu.Cashupno) inner join Ceremony cm on (cm.SYSIDCEREMONY = cum.Ceremonyno) where cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY <= <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:QtyGrads = SQL:Field1
        LOC:StageOrdered = SQL:Field2
        LOC:FamilyQtyA = SQL:Field3
        LOC:FamilyQtyB = SQL:Field4
        LOC:FamilyOrdered = (LOC:FamilyQtyA + LOC:FamilyQtyB)
        LOC:CashTotal = (SQL:Field11 + SQL:Field12 + SQL:Field13 + SQL:Field14 + SQL:Field15 + SQL:Field16)
        LOC:CreditTotalTotal = SQL:Field9
        LOC:ChequeTotal = SQL:Field8
        LOC:ElectronicTotal = SQL:Field10
        LOC:GrandTotal = SQL:Field6
        LOC:RationPerGrad = (SQL:Field5/SQL:Field7)
        LOC:StageOrderLine1Total = SQL:Field17
        LOC:StageOrderLine2Total = SQL:Field18
        LOC:StageOrderLine3Total = SQL:Field19
        LOC:StageOrderLine4Total = SQL:Field20
        LOC:PostalCalcTotal = SQL:Field22
        LOC:FamilyCalcTotal = SQL:Field21
        
        LOC:StageLine1Qty = SQL:Field24
        LOC:StageLine2Qty = SQL:Field25
        LOC:StageLine3Qty = SQL:Field26
        LOC:StageLine4Qty = SQL:Field27
    End

    ! Calculate Highres CD and Value for CDS using Same Criteria
    SQLScript{Prop:SQL}= 'Select Sum(cs.HighResFamilyCDOrdered), Sum(cs.HighResCDQty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) where cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:HighResFamilyCDOrdered = SQL:Field1
        LOC:HighresCDQty = SQL:Field2
    END
    
    SQLScript{Prop:SQL}= 'Select sum(ods.LINETOTALAMOUNT) From Cashup cu inner join CashupCeremony cum on (cum.CashupNo = cu.Cashupno) inner join Ceremony cm on (cm.SYSIDCEREMONY = cum.Ceremonyno) inner join Orders od on (od.SYSIDCEREMONY = cm.SYSIDCEREMONY) inner join OrderDetails ods on (ods.ORDERNO = od.ORDERNO) where cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY <= <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
            LOC:SumOrderedTotals = SQL:Field1
    END
        
    !INSA:SysIDInstitute = LOC:InstituteNo
    !Access:InstituteAlias.Fetch(INSA:PK_Institute)
        
    IF CLIp(LEFT(INSA:StagePackage1Description)) = '' then
        ?LOC:StageOrderLine1Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage1Description))
        ?LOC:StageOrderLine2Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage2Description))
        ?LOC:StageOrderLine3Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage3Description))
        ?LOC:StageOrderLine4Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage4Description))
    ELSE
        ?LOC:StageOrderLine1Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage1Description))
        ?LOC:StageOrderLine2Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage2Description))
        ?LOC:StageOrderLine3Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage3Description))
        ?LOC:StageOrderLine4Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage4Description))
    END

    Display()

    !ThisWindow.Reset

    SetCursor()

LocProc.ExecuteDateandInstituteQuery       Procedure()
    CODE

    SetCursor(Cursor:Wait)

    ! The process will calculate all the combined figures

    LOC:QueryType = 'Date and Institute Combined Query'

    Clear(CeremonyQueue)
    Free(CeremonyQueue)

    ! Calculate the Qty values
    SQLScript{Prop:SQL}= 'Select Sum(cs.NOOFGRADUATES), Sum(cs.STAGEODERQTY), Sum(cs.A4FAMILYQTY), Sum(cs.A3FAMILYQTY), Sum(cs.RATIONPERGRADUATE), sum(cs.GRANDTOTAL), Count(cs.CASHUPNO), Sum(cs.CHEQUES), Sum(cs.CREDITCARDS), Sum(cs.EFTAmount), Sum(cs.R200), Sum(cs.R100), Sum(cs.R50), Sum(cs.R20), Sum(cs.R10), Sum(cs.CHANGE), Sum(cs.StageLine1Value), Sum(cs.StageLine2Value), Sum(cs.StageLine3Value), Sum(cs.StageLine4Value), Sum(cs.TotalFamilyValue), Sum(cs.PostalValue), Sum(cs.HighResFamilyCDOrdered), Sum(cs.StageLine1Qty), sum(cs.StageLine2Qty), sum(cs.StageLine3Qty), sum(cs.StageLine4Qty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ' and cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:QtyGrads = SQL:Field1
        LOC:StageOrdered = SQL:Field2
        LOC:FamilyQtyA = SQL:Field3
        LOC:FamilyQtyB = SQL:Field4
        LOC:FamilyOrdered = (LOC:FamilyQtyA + LOC:FamilyQtyB)
        LOC:CashTotal = (SQL:Field11 + SQL:Field12 + SQL:Field13 + SQL:Field14 + SQL:Field15 + SQL:Field16)
        LOC:CreditTotalTotal = SQL:Field9
        LOC:ChequeTotal = SQL:Field8
        LOC:ElectronicTotal = SQL:Field10
        LOC:GrandTotal = SQL:Field6
        LOC:RationPerGrad = (SQL:Field5/SQL:Field7)
        LOC:StageOrderLine1Total = SQL:Field17
        LOC:StageOrderLine2Total = SQL:Field18
        LOC:StageOrderLine3Total = SQL:Field19
        LOC:StageOrderLine4Total = SQL:Field20
        LOC:PostalCalcTotal = SQL:Field22
        LOC:FamilyCalcTotal = SQL:Field21
        
        LOC:StageLine1Qty = SQL:Field24
        LOC:StageLine2Qty = SQL:Field25
        LOC:StageLine3Qty = SQL:Field26
        LOC:StageLine4Qty = SQL:Field27
    End

    ! Calculate Highres CD and Value for CDS using Same Criteria
    SQLScript{Prop:SQL}= 'Select Sum(cs.HighResFamilyCDOrdered), Sum(cs.HighResCDQty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ' and cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:HighResFamilyCDOrdered = SQL:Field1
        LOC:HighresCDQty = SQL:Field2
    END
       
    SQLScript{Prop:SQL}= 'Select sum(ods.LINETOTALAMOUNT) From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) inner join Orders od on (od.SYSIDCEREMONY = cm.SYSIDCEREMONY) inner join OrderDetails ods on (ods.ORDERNO = od.ORDERNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ' and cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
            LOC:SumOrderedTotals = SQL:Field1
    END
        
       
    INSA:SysIDInstitute = LOC:InstituteNo
    !Access:InstituteAlias.Fetch(INSA:PK_Institute)
        
    IF CLIp(LEFT(INSA:StagePackage1Description)) = '' then
        ?LOC:StageOrderLine1Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage1Description))
        ?LOC:StageOrderLine2Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage2Description))
        ?LOC:StageOrderLine3Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage3Description))
        ?LOC:StageOrderLine4Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage4Description))
    ELSE
        ?LOC:StageOrderLine1Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage1Description))
        ?LOC:StageOrderLine2Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage2Description))
        ?LOC:StageOrderLine3Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage3Description))
        ?LOC:StageOrderLine4Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage4Description))
    END
        
        
    Display()

    !ThisWindow.Reset

    SetCursor()

LocProc.ExecuteInstituteQuery       Procedure()
    CODE

    SetCursor(Cursor:Wait)

    ! The process will calculate all the combined figures

    LOC:QueryType = 'Institute Query'

    ! Calculate the Qty values
    SQLScript{Prop:SQL}= 'Select Sum(cs.NOOFGRADUATES), Sum(cs.STAGEODERQTY), Sum(cs.A4FAMILYQTY), Sum(cs.A3FAMILYQTY), Sum(cs.RATIONPERGRADUATE), sum(cs.GRANDTOTAL), Count(cs.CASHUPNO), Sum(cs.CHEQUES), Sum(cs.CREDITCARDS), Sum(cs.EFTAmount), Sum(cs.R200), Sum(cs.R100), Sum(cs.R50), Sum(cs.R20), Sum(cs.R10), Sum(cs.CHANGE), Sum(cs.StageLine1Value), Sum(cs.StageLine2Value), Sum(cs.StageLine3Value), Sum(cs.StageLine4Value), Sum(cs.TotalFamilyValue), Sum(cs.PostalValue), Sum(cs.HighResFamilyCDOrdered), Sum(cs.StageLine1Qty), sum(cs.StageLine2Qty), sum(cs.StageLine3Qty), sum(cs.StageLine4Qty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:QtyGrads = SQL:Field1
        LOC:StageOrdered = SQL:Field2
        LOC:FamilyQtyA = SQL:Field3
        LOC:FamilyQtyB = SQL:Field4
        LOC:FamilyOrdered = (LOC:FamilyQtyA + LOC:FamilyQtyB)
        LOC:CashTotal = (SQL:Field11 + SQL:Field12 + SQL:Field13 + SQL:Field14 + SQL:Field15 + SQL:Field16)
        LOC:CreditTotalTotal = SQL:Field9
        LOC:ChequeTotal = SQL:Field8
        LOC:ElectronicTotal = SQL:Field10
        LOC:GrandTotal = SQL:Field6
        LOC:RationPerGrad = (SQL:Field5/SQL:Field7)
        LOC:StageOrderLine1Total = SQL:Field17
        LOC:StageOrderLine2Total = SQL:Field18
        LOC:StageOrderLine3Total = SQL:Field19
        LOC:StageOrderLine4Total = SQL:Field20
        LOC:PostalCalcTotal = SQL:Field22
        LOC:FamilyCalcTotal = SQL:Field21
        
        LOC:StageLine1Qty = SQL:Field24
        LOC:StageLine2Qty = SQL:Field25
        LOC:StageLine3Qty = SQL:Field26
        LOC:StageLine4Qty = SQL:Field27        
    End

    ! Calculate Highres CD and Value for CDS using Same Criteria
    SQLScript{Prop:SQL}= 'Select Sum(cs.HighResFamilyCDOrdered), Sum(cs.HighResCDQty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:HighResFamilyCDOrdered = SQL:Field1
        LOC:HighresCDQty = SQL:Field2
    END

    SQLScript{Prop:SQL}= 'Select sum(ods.LINETOTALAMOUNT) From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) inner join Orders od on (od.SYSIDCEREMONY = cm.SYSIDCEREMONY) inner join OrderDetails ods on (ods.ORDERNO = od.ORDERNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ')'
    Next(SQLScript)
    If Not Error() then
            LOC:SumOrderedTotals = SQL:Field1
    END
        
    INSA:SysIDInstitute = LOC:InstituteNo
    !Access:InstituteAlias.Fetch(INSA:PK_Institute)
        
    IF CLIp(LEFT(INSA:StagePackage1Description)) = '' then
        ?LOC:StageOrderLine1Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage1Description))
        ?LOC:StageOrderLine2Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage2Description))
        ?LOC:StageOrderLine3Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage3Description))
        ?LOC:StageOrderLine4Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage4Description))
    ELSE
        ?LOC:StageOrderLine1Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage1Description))
        ?LOC:StageOrderLine2Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage2Description))
        ?LOC:StageOrderLine3Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage3Description))
        ?LOC:StageOrderLine4Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage4Description))
    END
                
        
    Display()

    !ThisWindow.Reset

    SetCursor()

LocProc.ExecuteAllCombinedQuery       Procedure()
    CODE

    SetCursor(Cursor:Wait)

    ! The process will calculate all the combined figures

    LOC:QueryType = 'Combined Query'

    ! Calculate the Qty values
    SQLScript{Prop:SQL}= 'Select Sum(cs.NOOFGRADUATES), Sum(cs.STAGEODERQTY), Sum(cs.A4FAMILYQTY), Sum(cs.A3FAMILYQTY), Sum(cs.RATIONPERGRADUATE), sum(cs.GRANDTOTAL), Count(cs.CASHUPNO), Sum(cs.CHEQUES), Sum(cs.CREDITCARDS), Sum(cs.EFTAmount), Sum(cs.R200), Sum(cs.R100), Sum(cs.R50), Sum(cs.R20), Sum(cs.R10), Sum(cs.CHANGE), Sum(cs.StageLine1Value), Sum(cs.StageLine2Value), Sum(cs.StageLine3Value), Sum(cs.StageLine4Value), Sum(cs.TotalFamilyValue), Sum(cs.PostalValue), Sum(cs.HighResFamilyCDOrdered), Sum(cs.StageLine1Qty), sum(cs.StageLine2Qty), sum(cs.StageLine3Qty), sum(cs.StageLine4Qty) From Cashup cs'
    Next(SQLScript)
    If Not Error() then
        LOC:QtyGrads = SQL:Field1
        LOC:StageOrdered = SQL:Field2
        LOC:FamilyQtyA = SQL:Field3
        LOC:FamilyQtyB = SQL:Field4
        LOC:FamilyOrdered = (LOC:FamilyQtyA + LOC:FamilyQtyB)
        LOC:CashTotal = (SQL:Field11 + SQL:Field12 + SQL:Field13 + SQL:Field14 + SQL:Field15 + SQL:Field16)
        LOC:CreditTotalTotal = SQL:Field9
        LOC:ChequeTotal = SQL:Field8
        LOC:ElectronicTotal = SQL:Field10
        LOC:GrandTotal = SQL:Field6
        LOC:RationPerGrad = (SQL:Field5/SQL:Field7)
        LOC:StageOrderLine1Total = SQL:Field17
        LOC:StageOrderLine2Total = SQL:Field18
        LOC:StageOrderLine3Total = SQL:Field19
        LOC:StageOrderLine4Total = SQL:Field20
        LOC:PostalCalcTotal = SQL:Field22
        LOC:FamilyCalcTotal = SQL:Field21

        LOC:StageLine1Qty = SQL:Field24
        LOC:StageLine2Qty = SQL:Field25
        LOC:StageLine3Qty = SQL:Field26
        LOC:StageLine4Qty = SQL:Field27        
    End
        
    ! Calculate Highres CD and Value for CDS using Same Criteria
    SQLScript{Prop:SQL}= 'Select Sum(cs.HighResFamilyCDOrdered), Sum(cs.HighResCDQty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO))'
    Next(SQLScript)
    If Not Error() then
        LOC:HighResFamilyCDOrdered = SQL:Field1
        LOC:HighresCDQty = SQL:Field2
    END

    LOC:SumOrderedTotals = -999999
        
    INSA:SysIDInstitute = LOC:InstituteNo
    !Access:InstituteAlias.Fetch(INSA:PK_Institute)
        
    IF CLIp(LEFT(INSA:StagePackage1Description)) = '' then
        ?LOC:StageOrderLine1Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage1Description))
        ?LOC:StageOrderLine2Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage2Description))
        ?LOC:StageOrderLine3Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage3Description))
        ?LOC:StageOrderLine4Total:Prompt{PROP:Text} = CLIP(LEFT(GLO:StagePackage4Description))
    ELSE
        ?LOC:StageOrderLine1Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage1Description))
        ?LOC:StageOrderLine2Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage2Description))
        ?LOC:StageOrderLine3Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage3Description))
        ?LOC:StageOrderLine4Total:Prompt{PROP:Text} = CLIP(LEFT(INSA:StagePackage4Description))
    END
        
    Display()

    !ThisWindow.Reset

    SetCursor()

LocProc.RefreshScreens  Procedure()
    CODE

    ! Calculate the Qty of Print Jobs
    SQLScript{Prop:SQL} = 'Select Count(PrintServiceDetailNo) From PrintServiceDetail where Status = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:PrintServiceCount = SQL:Field1
    End

    ! Calculate the Qty of SMS messages to be sent
    SQLScript{Prop:SQL} = 'Select Count(CommunicationNo) From Communication where Status = 0 and OutputType = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:SMSCount = SQL:Field1
    End

    ! Calculate the Qty of Email messages to be sent
    SQLScript{Prop:SQL} = 'Select Count(CommunicationNo) From Communication where Status = 0 and OutputType = 1'
    Next(SQLScript)
    If Not Error() then
        LOC:EmailCount = SQL:Field1
    End

    !BRW8.ResetFromFile
    !BRW8.ResetSort(True)

    !BRW7.ResetFromFile
    !BRW7.ResetSort(True)

    ! Calculate the Overwiew Details
    SQLScript{Prop:SQL} = 'Select Count(ClientNo) From Client where Status = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfOrdersClients = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfNew = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 1'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfInProcess = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 5'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfPosted = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 6'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfCancelled = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 7'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfCollected = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where SysIdOrderStatus = 8'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfReturned = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcards = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 0'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsOpen = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 1'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsLab = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 2'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsCompleted = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 3'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsCollected = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 4'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsPosted = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(JobcardNo) From Jobcard where Status = 5'
    Next(SQLScript)
    If Not Error() then
        LOC:NoOfJobcardsReturned = SQL:Field1
    End

    SQLScript{Prop:SQL} = 'Select Count(OrderNo) From Orders where NedbankBoardCheck = 1'
    Next(SQLScript)
    If Not Error() then
        LOC:CountofBankBoards = SQL:Field1
    End

    LOC:LastUpdateString = 'Last Refresh date and time: ' & format(Today(),@d06b) & ' ' & Format(Clock(),@t01b)

    GLO:MessageCredit = GetConfig('GLO:MessageCredit','0')

LocProc.ClearAll      Procedure()
    CODE

    ! Clear all the local variables
    LOC:QtyGrads = 0
    LOC:StageOrdered = 0
    LOC:FamilyQtyA = 0
    LOC:FamilyQtyB = 0
    LOC:FamilyOrdered = 0
    LOC:CashTotal = 0
    LOC:CreditTotalTotal = 0
    LOC:ChequeTotal = 0
    LOC:ElectronicTotal = 0
    LOC:GrandTotal = 0
    LOC:RationPerGrad = 0

    LOC:FamilyCalcTotal = 0
    LOC:PostalCalcTotal = 0
    LOC:HighresCDQty = 0
	LOC:HighResFamilyCDOrdered = 0
	LOC:CountedPhotostaken = 0
        
    ! Clear the filter
    GLO:DashboardQueryType = 0

    FromDate = Today()
    ToDate = Today()

    LOC:InstituteNo = 0
    LOC:InstituteDescription = ''

    LOC:LastUpdateString = 'Cleared the results'

    Display()

    !ThisWindow.Reset

LocProc.ClearAll2016      Procedure()
    CODE

    ! Clear all the local variables
    LOC:QtyGrads = 0
    LOC:StageOrdered = 0
    LOC:FamilyQtyA = 0
    LOC:FamilyQtyB = 0
    LOC:FamilyOrdered = 0
    LOC:CashTotal = 0
    LOC:CreditTotalTotal = 0
    LOC:ChequeTotal = 0
    LOC:ElectronicTotal = 0
    LOC:GrandTotal = 0
    LOC:RationPerGrad = 0

    LOC:FamilyCalcTotal = 0
    LOC:PostalCalcTotal = 0
    LOC:HighresCDQty = 0
    LOC:HighResFamilyCDOrdered = 0
	LOC:CountedPhotostaken = 0
		
    ! Clear the filter
    GLO:DashboardQueryType = 0

    FromDate = Today()
    ToDate = Today()

    LOC:InstituteNo = 0
    LOC:InstituteDescription = ''

    LOC:LastUpdateString = 'Cleared the results'

    stream(Dashboard2016Alias)
    Dasha:Dashboard2016No = 1
    Set(Dasha:Dashboard2016_PK,Dasha:Dashboard2016_PK)
    LOOP
        !if Access:Dashboard2016Alias.Next() then break.
        !Access:Dashboard2016Alias.DeleteRecord(0)
    END

    FLUSH(Dashboard2016Alias)
        
    !BRW17.ResetFromFile()
    !BRW17.ResetSort(True)
      
        
    Display()

    !ThisWindow.Reset

LocProc.CalculateQtyViaSize Procedure()
    CODE

    Disable(?Button13)
    Update()

    Clear(SizeQtyQueue)
    Free(SizeQtyQueue)

    ! Loop through all the sized
    LPHOSIZ:SysIdPhotoSize = 1
    Set(LPHOSIZ:PhotoSize_PK,LPHOSIZ:PhotoSize_PK)
    Loop
        !if Access:LPhotosize.Next() then break.

        Clear(SizeQtyQueue)

        SQQ:SizeDescription = LPHOSIZ:DescriptionPhotoSize

        LOC:CurrentAction = 'Calculating qty for size: ' & LPHOSIZ:DescriptionPhotoSize
        Display()

        Update

        Case LOC:Type
        of 0
            ! Calculate for Grads
            SQLScript{Prop:SQL} = 'Select Sum(ods.QUANTITY) from Orders od inner join OrderDetails ods on (od.OrderNo = ods.OrderNo) where od.DATECREATE >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and od.DATECREATE < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39> and ods.Quantity > 0 and ods.SYSIDPHOTOSIZE = ' & LPHOSIZ:SysIdPhotoSize
            Next(SQLScript)
            If Not Error() then
                SQQ:Qty = SQL:Field1
            End
        of 1
            ! Calculate for Jobcards
            SQLScript{Prop:SQL} = 'Select Sum(ji.QTY) from JobCardItem ji inner join PackageDetail pd on (pd.PACKAGEDETAILNO = ji.PACKAGEDETAILNO) where ji.JOBCARDNO in (Select jb.JOBCARDNO from Jobcard jb where jb.BEGINDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and jb.BEGINDATETIME < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39> ' & ') and pd.PHOTOSIZENO = ' & LPHOSIZ:SysIdPhotoSize
            Next(SQLScript)
            If Not Error() then
                SQQ:Qty = SQQ:Qty + SQL:Field1
            End
        of 2
            ! Calculate for Schools
            !SQQ:Qty = SQQ:Qty +
        END

        Add(SizeQtyQueue)
    END

    LOC:CurrentAction = 'Calculation Completed'

    Enable(?Button13)
    Display()
    Update()

LocProc.ExecuteFromAndToDateQuery2016       Procedure()
    CODE

    SetCursor(Cursor:Wait)

    ! The process will calculate all the combined figures

    LOC:QueryType = 'Combined Query'

    ! Calculate the Qty values
    SQLScript{Prop:SQL}= 'Select Sum(cs.NOOFGRADUATES), Sum(cs.STAGEODERQTY), Sum(cs.A4FAMILYQTY), Sum(cs.A3FAMILYQTY), Sum(cs.RATIONPERGRADUATE), sum(cs.GRANDTOTAL), Count(cs.CASHUPNO), Sum(cs.CHEQUES), Sum(cs.CREDITCARDS), Sum(cs.EFTAmount), Sum(cs.R200), Sum(cs.R100), Sum(cs.R50), Sum(cs.R20), Sum(cs.R10), Sum(cs.CHANGE), Sum(cs.StageLine1Value), Sum(cs.StageLine2Value), Sum(cs.StageLine3Value), Sum(cs.StageLine4Value), Sum(cs.TotalFamilyValue), Sum(cs.PostalValue), Sum(cs.HighResFamilyCDOrdered), Sum(cs.StageLine1Qty), sum(cs.StageLine2Qty), sum(cs.StageLine3Qty), sum(cs.StageLine4Qty) From Cashup cs where cs.cashupno in (Select distinct(cu.cashupno) From Cashup cu inner join CashupCeremony cum on (cum.CashupNo = cu.Cashupno) inner join Ceremony cm on (cm.SYSIDCEREMONY = cum.Ceremonyno) where cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY <= <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:QtyGrads = SQL:Field1
        LOC:StageOrdered = SQL:Field2
        LOC:FamilyQtyA = SQL:Field3
        LOC:FamilyQtyB = SQL:Field4
        LOC:FamilyOrdered = (LOC:FamilyQtyA + LOC:FamilyQtyB)
        LOC:CashTotal = (SQL:Field11 + SQL:Field12 + SQL:Field13 + SQL:Field14 + SQL:Field15 + SQL:Field16)
        LOC:CreditTotalTotal = SQL:Field9
        LOC:ChequeTotal = SQL:Field8
        LOC:ElectronicTotal = SQL:Field10
        LOC:GrandTotal = SQL:Field6
        LOC:RationPerGrad = (SQL:Field5/SQL:Field7)
        LOC:PostalCalcTotal = SQL:Field22
        LOC:FamilyCalcTotal = SQL:Field21
    End

    ! Calculate Highres CD and Value for CDS using Same Criteria
    SQLScript{Prop:SQL}= 'Select Sum(cs.HighResFamilyCDOrdered), Sum(cs.HighResCDQty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) where cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:HighResFamilyCDOrdered = SQL:Field1
        LOC:HighresCDQty = SQL:Field2
    END
    
    SQLScript{Prop:SQL}= 'Select sum(ods.LINETOTALAMOUNT) From Cashup cu inner join CashupCeremony cum on (cum.CashupNo = cu.Cashupno) inner join Ceremony cm on (cm.SYSIDCEREMONY = cum.Ceremonyno) inner join Orders od on (od.SYSIDCEREMONY = cm.SYSIDCEREMONY) inner join OrderDetails ods on (ods.ORDERNO = od.ORDERNO) where cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY <= <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
            LOC:SumOrderedTotals = SQL:Field1
    END
        
    INSA:SysIDInstitute = LOC:InstituteNo
    !Access:InstituteAlias.Fetch(INSA:PK_Institute)

    stream(Dashboard2016Alias)
    Dasha:Dashboard2016No = 1
    Set(Dasha:Dashboard2016_PK,Dasha:Dashboard2016_PK)
    LOOP
        !if Access:Dashboard2016Alias.Next() then break.
        !Access:Dashboard2016Alias.DeleteRecord(0)
    END

    FLUSH(Dashboard2016Alias)
        
    !BRW17.ResetFromFile()
    !BRW17.ResetSort(True)
        
    
    !Now calculate the cashuppackags
        SQLScript2{Prop:SQL}= 'select cp.PackageDescription, PackageNo, sum(cp.TotalQty) , Sum(cp.TotalValue) from CashupPackages cp where cp.CashupNo in (Select distinct(cu.cashupno) From Cashup cu inner join CashupCeremony cum on (cum.CashupNo = cu.Cashupno) inner join Ceremony cm on (cm.SYSIDCEREMONY = cum.Ceremonyno) where cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY <= <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ') group by cp.PackageDescription, cp.PackageNo'
    Loop xcv# = 1 to 100
        Next(SQLScript2)
        
        If Not Error() then
         
        Dash:PackageNo = SQLS:Field2
        Dash:PackageDescription = SQLS:Field1
        Dash:TotalQty = SQLS:Field3
        Dash:TotalValue = SQLS:Field4
        !Access:Dashboard2016.Insert()
        END
    END
    
    !BRW17.ResetFromFile()
	!BRW17.ResetSort(True)        
		
	LocProc.LoadPhotosTakenDate()
		
    Display()

    !ThisWindow.Reset

    SetCursor()

LocProc.ExecuteAllCombinedQuery2016       Procedure()
    CODE

    SetCursor(Cursor:Wait)

    ! The process will calculate all the combined figures

    LOC:QueryType = 'Combined Query'

    ! Calculate the Qty values
    SQLScript{Prop:SQL}= 'Select Sum(cs.NOOFGRADUATES), Sum(cs.STAGEODERQTY), Sum(cs.A4FAMILYQTY), Sum(cs.A3FAMILYQTY), Sum(cs.RATIONPERGRADUATE), sum(cs.GRANDTOTAL), Count(cs.CASHUPNO), Sum(cs.CHEQUES), Sum(cs.CREDITCARDS), Sum(cs.EFTAmount), Sum(cs.R200), Sum(cs.R100), Sum(cs.R50), Sum(cs.R20), Sum(cs.R10), Sum(cs.CHANGE), Sum(cs.StageLine1Value), Sum(cs.StageLine2Value), Sum(cs.StageLine3Value), Sum(cs.StageLine4Value), Sum(cs.TotalFamilyValue), Sum(cs.PostalValue), Sum(cs.HighResFamilyCDOrdered), Sum(cs.StageLine1Qty), sum(cs.StageLine2Qty), sum(cs.StageLine3Qty), sum(cs.StageLine4Qty) From Cashup cs'
    Next(SQLScript)
    If Not Error() then
        LOC:QtyGrads = SQL:Field1
        LOC:StageOrdered = SQL:Field2
        LOC:FamilyQtyA = SQL:Field3
        LOC:FamilyQtyB = SQL:Field4
        LOC:FamilyOrdered = (LOC:FamilyQtyA + LOC:FamilyQtyB)
        LOC:CashTotal = (SQL:Field11 + SQL:Field12 + SQL:Field13 + SQL:Field14 + SQL:Field15 + SQL:Field16)
        LOC:CreditTotalTotal = SQL:Field9
        LOC:ChequeTotal = SQL:Field8
        LOC:ElectronicTotal = SQL:Field10
        LOC:GrandTotal = SQL:Field6
        LOC:RationPerGrad = (SQL:Field5/SQL:Field7)
        LOC:PostalCalcTotal = SQL:Field22
        LOC:FamilyCalcTotal = SQL:Field21
    END
        
    stream(Dashboard2016Alias)
    Dasha:Dashboard2016No = 1
    Set(Dasha:Dashboard2016_PK,Dasha:Dashboard2016_PK)
    LOOP
        !if Access:Dashboard2016Alias.Next() then break.
        !Access:Dashboard2016Alias.DeleteRecord(0)
    END

    FLUSH(Dashboard2016Alias)
        
    !BRW17.ResetFromFile()
    !BRW17.ResetSort(True)      
    
    !Now calculate the cashuppackags
        SQLScript2{Prop:SQL}= 'select PackageDescription, PackageNo, sum(TotalQty) , Sum(TotalValue) from CashupPackages group by PackageDescription, PackageNo'
    Loop xcv# = 1 to 100
        Next(SQLScript2)
        
        If Not Error() then
         
        Dash:PackageNo = SQLS:Field2
        Dash:PackageDescription = SQLS:Field1
        Dash:TotalQty = SQLS:Field3
        Dash:TotalValue = SQLS:Field4
        !Access:Dashboard2016.Insert()
        END
    END
    
    !BRW17.ResetFromFile()
    !BRW17.ResetSort(True)
        
    ! Calculate Highres CD and Value for CDS using Same Criteria
    SQLScript{Prop:SQL}= 'Select Sum(cs.HighResFamilyCDOrdered), Sum(cs.HighResCDQty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO))'
    Next(SQLScript)
    If Not Error() then
        LOC:HighResFamilyCDOrdered = SQL:Field1
        LOC:HighresCDQty = SQL:Field2
    END

    LOC:SumOrderedTotals = -999999
        
    INSA:SysIDInstitute = LOC:InstituteNo
    !Access:InstituteAlias.Fetch(INSA:PK_Institute)
        
    Display()

    !BRW17.ResetFromFile()
    !BRW17.ResetSort(True)

	Display()
    !ThisWindow.Reset

    SetCursor()

LocProc.ExecuteDateandInstituteQuery2016       Procedure()
    CODE

    SetCursor(Cursor:Wait)

    ! The process will calculate all the combined figures

    LOC:QueryType = 'Date and Institute Combined Query'

    Clear(CeremonyQueue)
    Free(CeremonyQueue)

    ! Calculate the Qty values
    SQLScript{Prop:SQL}= 'Select Sum(cs.NOOFGRADUATES), Sum(cs.STAGEODERQTY), Sum(cs.A4FAMILYQTY), Sum(cs.A3FAMILYQTY), Sum(cs.RATIONPERGRADUATE), sum(cs.GRANDTOTAL), Count(cs.CASHUPNO), Sum(cs.CHEQUES), Sum(cs.CREDITCARDS), Sum(cs.EFTAmount), Sum(cs.R200), Sum(cs.R100), Sum(cs.R50), Sum(cs.R20), Sum(cs.R10), Sum(cs.CHANGE), Sum(cs.StageLine1Value), Sum(cs.StageLine2Value), Sum(cs.StageLine3Value), Sum(cs.StageLine4Value), Sum(cs.TotalFamilyValue), Sum(cs.PostalValue), Sum(cs.HighResFamilyCDOrdered), Sum(cs.StageLine1Qty), sum(cs.StageLine2Qty), sum(cs.StageLine3Qty), sum(cs.StageLine4Qty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ' and cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:QtyGrads = SQL:Field1
        LOC:StageOrdered = SQL:Field2
        LOC:FamilyQtyA = SQL:Field3
        LOC:FamilyQtyB = SQL:Field4
        LOC:FamilyOrdered = (LOC:FamilyQtyA + LOC:FamilyQtyB)
        LOC:CashTotal = (SQL:Field11 + SQL:Field12 + SQL:Field13 + SQL:Field14 + SQL:Field15 + SQL:Field16)
        LOC:CreditTotalTotal = SQL:Field9
        LOC:ChequeTotal = SQL:Field8
        LOC:ElectronicTotal = SQL:Field10
        LOC:GrandTotal = SQL:Field6
        LOC:RationPerGrad = (SQL:Field5/SQL:Field7)
        LOC:PostalCalcTotal = SQL:Field22
        LOC:FamilyCalcTotal = SQL:Field21
    End
        
    stream(Dashboard2016Alias)
    Dasha:Dashboard2016No = 1
    Set(Dasha:Dashboard2016_PK,Dasha:Dashboard2016_PK)
    LOOP
        !if Access:Dashboard2016Alias.Next() then break.
        !Access:Dashboard2016Alias.DeleteRecord(0)
    END

    FLUSH(Dashboard2016Alias)
        
    !BRW17.ResetFromFile()
    !BRW17.ResetSort(True)
      
    
    !Now calculate the cashuppackags
    SQLScript2{Prop:SQL}= 'select cp.PackageDescription, PackageNo, sum(cp.TotalQty) , Sum(cp.TotalValue) from CashupPackages cp where cp.CashupNo in (Select distinct(cu.cashupno) From Cashup cu inner join CashupCeremony cum on (cum.CashupNo = cu.Cashupno) inner join Ceremony cm on (cm.SYSIDCEREMONY = cum.Ceremonyno) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ' and cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY <= <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ') group by cp.PackageDescription, cp.PackageNo'
    SETCLIPBOARD('select cp.PackageDescription, PackageNo, sum(cp.TotalQty) , Sum(cp.TotalValue) from CashupPackages cp where cp.CashupNo in (Select distinct(cu.cashupno) From Cashup cu inner join CashupCeremony cum on (cum.CashupNo = cu.Cashupno) inner join Ceremony cm on (cm.SYSIDCEREMONY = cum.Ceremonyno) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ' and cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY <= <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ') group by cp.PackageDescription, cp.PackageNo')
    Loop xcv# = 1 to 100
        Next(SQLScript2)
        
        If Not Error() then
         
        Dash:PackageNo = SQLS:Field2
        Dash:PackageDescription = SQLS:Field1
        Dash:TotalQty = SQLS:Field3
        Dash:TotalValue = SQLS:Field4
        !Access:Dashboard2016.Insert()
        END
    END        

    ! Calculate Highres CD and Value for CDS using Same Criteria
    SQLScript{Prop:SQL}= 'Select Sum(cs.HighResFamilyCDOrdered), Sum(cs.HighResCDQty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ' and cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:HighResFamilyCDOrdered = SQL:Field1
        LOC:HighresCDQty = SQL:Field2
    END
       
    SQLScript{Prop:SQL}= 'Select sum(ods.LINETOTALAMOUNT) From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) inner join Orders od on (od.SYSIDCEREMONY = cm.SYSIDCEREMONY) inner join OrderDetails ods on (ods.ORDERNO = od.ORDERNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ' and cm.DATECEREMONY >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and cm.DATECEREMONY < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ')'
    Next(SQLScript)
    If Not Error() then
            LOC:SumOrderedTotals = SQL:Field1
    END
        
        
    INSA:SysIDInstitute = LOC:InstituteNo
    !Access:InstituteAlias.Fetch(INSA:PK_Institute)
        
    Display()

    !BRW17.ResetFromFile()
    !BRW17.ResetSort(True)

	LocProc.LoadPhotosTakenDateInstitute()		

    !ThisWindow.Reset

    SetCursor()

LocProc.ExecuteInstituteQuery2016       Procedure()
    CODE

    SetCursor(Cursor:Wait)

    ! The process will calculate all the combined figures

    LOC:QueryType = 'Institute Query'

    ! Calculate the Qty values
    SQLScript{Prop:SQL}= 'Select Sum(cs.NOOFGRADUATES), Sum(cs.STAGEODERQTY), Sum(cs.A4FAMILYQTY), Sum(cs.A3FAMILYQTY), Sum(cs.RATIONPERGRADUATE), sum(cs.GRANDTOTAL), Count(cs.CASHUPNO), Sum(cs.CHEQUES), Sum(cs.CREDITCARDS), Sum(cs.EFTAmount), Sum(cs.R200), Sum(cs.R100), Sum(cs.R50), Sum(cs.R20), Sum(cs.R10), Sum(cs.CHANGE), Sum(cs.StageLine1Value), Sum(cs.StageLine2Value), Sum(cs.StageLine3Value), Sum(cs.StageLine4Value), Sum(cs.TotalFamilyValue), Sum(cs.PostalValue), Sum(cs.HighResFamilyCDOrdered), Sum(cs.StageLine1Qty), sum(cs.StageLine2Qty), sum(cs.StageLine3Qty), sum(cs.StageLine4Qty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:QtyGrads = SQL:Field1
        LOC:StageOrdered = SQL:Field2
        LOC:FamilyQtyA = SQL:Field3
        LOC:FamilyQtyB = SQL:Field4
        LOC:FamilyOrdered = (LOC:FamilyQtyA + LOC:FamilyQtyB)
        LOC:CashTotal = (SQL:Field11 + SQL:Field12 + SQL:Field13 + SQL:Field14 + SQL:Field15 + SQL:Field16)
        LOC:CreditTotalTotal = SQL:Field9
        LOC:ChequeTotal = SQL:Field8
        LOC:ElectronicTotal = SQL:Field10
        LOC:GrandTotal = SQL:Field6
        LOC:RationPerGrad = (SQL:Field5/SQL:Field7)
        LOC:PostalCalcTotal = SQL:Field22
        LOC:FamilyCalcTotal = SQL:Field21
    End

    stream(Dashboard2016Alias)
    Dasha:Dashboard2016No = 1
    Set(Dasha:Dashboard2016_PK,Dasha:Dashboard2016_PK)
    LOOP
        !if Access:Dashboard2016Alias.Next() then break.
        !Access:Dashboard2016Alias.DeleteRecord(0)
    END

    FLUSH(Dashboard2016Alias)
        
    !BRW17.ResetFromFile()
    !BRW17.ResetSort(True)
        
    !Now calculate the cashuppackags
        SQLScript2{Prop:SQL}= 'select cp.PackageDescription, PackageNo, sum(cp.TotalQty) , Sum(cp.TotalValue) from CashupPackages cp where cp.CashupNo in (Select distinct(cu.cashupno) From Cashup cu inner join CashupCeremony cum on (cum.CashupNo = cu.Cashupno) inner join Ceremony cm on (cm.SYSIDCEREMONY = cum.Ceremonyno) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ') group by cp.PackageDescription, cp.PackageNo'
        SETCLIPBOARD('select cp.PackageDescription, PackageNo, sum(cp.TotalQty) , Sum(cp.TotalValue) from CashupPackages cp where cp.CashupNo in (Select distinct(cu.cashupno) From Cashup cu inner join CashupCeremony cum on (cum.CashupNo = cu.Cashupno) inner join Ceremony cm on (cm.SYSIDCEREMONY = cum.Ceremonyno) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ') group by cp.PackageDescription, cp.PackageNo')
    Loop xcv# = 1 to 100
        Next(SQLScript2)
        
        If Not Error() then
         
        Dash:PackageNo = SQLS:Field2
        Dash:PackageDescription = SQLS:Field1
        Dash:TotalQty = SQLS:Field3
        Dash:TotalValue = SQLS:Field4
        !Access:Dashboard2016.Insert()
        END
    END                
        
    ! Calculate Highres CD and Value for CDS using Same Criteria
    SQLScript{Prop:SQL}= 'Select Sum(cs.HighResFamilyCDOrdered), Sum(cs.HighResCDQty) From Cashup cs where cs.cashupno in (Select Distinct cu.cashupNo From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ')'
    Next(SQLScript)
    If Not Error() then
        LOC:HighResFamilyCDOrdered = SQL:Field1
        LOC:HighresCDQty = SQL:Field2
    END

    SQLScript{Prop:SQL}= 'Select sum(ods.LINETOTALAMOUNT) From Cashup cu inner join cashupCeremony cc on (cc.Cashupno = cu.cashupNo) inner join Ceremony cm on (cm.SysIDCeremony = cc.CEREMONYNO) inner join Orders od on (od.SYSIDCEREMONY = cm.SYSIDCEREMONY) inner join OrderDetails ods on (ods.ORDERNO = od.ORDERNO) where cm.SYSIDINSTITUTE = ' & LOC:InstituteNo & ')'
    Next(SQLScript)
    If Not Error() then
            LOC:SumOrderedTotals = SQL:Field1
    END
        
    INSA:SysIDInstitute = LOC:InstituteNo
    !Access:InstituteAlias.Fetch(INSA:PK_Institute)
           
    Display()

    !BRW17.ResetFromFile()
    !BRW17.ResetSort(True)
		
	LocProc.LoadPhotosTakenInstitute()		
		
	Display()		
		
    !ThisWindow.Reset

    SetCursor()
        
LocProc.ExecuteJobcardQueryPreSchool       Procedure()
    CODE

    SetCursor(Cursor:Wait)

    ! The process will calculate all the jobcard figures
    LOC:QueryType = 'Jobcard Query'

    stream(Dashboard2016Alias)
    Dasha:Dashboard2016No = 1
    Set(Dasha:Dashboard2016_PK,Dasha:Dashboard2016_PK)
    LOOP
        !if Access:Dashboard2016Alias.Next() then break.
        !Access:Dashboard2016Alias.DeleteRecord(0)
    END

    FLUSH(Dashboard2016Alias)
        
    !BRW21.ResetFromFile()
    !BRW21.ResetSort(True)

    !Now calculate the school
    SQLScript2{Prop:SQL}= 'Select PACKAGEDETAILNO, JOBCARDITEMDESCRIPTION, QTY, TOTAL from JobCardItem where JOBCARDNO in (Select JOBCARDNO from Jobcard where BEGINDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and BEGINDATETIME < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ' and DEPARTMENTNO = 8)'
    SETCLIPBOARD('Select PACKAGEDETAILNO, JOBCARDITEMDESCRIPTION, QTY, TOTAL from JobCardItem where JOBCARDNO in (Select JOBCARDNO from Jobcard where BEGINDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and BEGINDATETIME < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ' and DEPARTMENTNO = 8)')
    Loop xcv# = 1 to 1000
        Next(SQLScript2)
        
        If Not Error() then
            Dash:PackageNo = SQLS:Field1
            Dash:PackageDescription = SQLS:Field2
            Dash:TotalQty = SQLS:Field3
            Dash:TotalValue = SQLS:Field4
            !Access:Dashboard2016.Insert()
        END
    END                
        
        SQLScript{Prop:SQL}=  'Select Sum(TOTAL) from JobCardItem where JOBCARDNO in (Select JOBCARDNO from Jobcard where BEGINDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and BEGINDATETIME < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ' and DEPARTMENTNO = 8)'
        SETCLIPBOARD('Select Sum(TOTAL) from JobCardItem where JOBCARDNO in (Select JOBCARDNO from Jobcard where BEGINDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and BEGINDATETIME < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ' and DEPARTMENTNO = 8)')
        
    Next(SQLScript)
    If Not Error() then
        LOC:SchoolTotal = SQL:Field1
    END
           
    Display()

    !BRW21.ResetFromFile()
    !BRW21.ResetSort(True)

    !ThisWindow.Reset

    SetCursor()

LocProc.ExecuteJobcardQuerySchool       Procedure()
    CODE

    SetCursor(Cursor:Wait)

    ! The process will calculate all the jobcard figures
    LOC:QueryType = 'Jobcard Query'

    stream(Dashboard2016Alias)
    Dasha:Dashboard2016No = 1
    Set(Dasha:Dashboard2016_PK,Dasha:Dashboard2016_PK)
    LOOP
        !if Access:Dashboard2016Alias.Next() then break.
        !Access:Dashboard2016Alias.DeleteRecord(0)
    END

    FLUSH(Dashboard2016Alias)
        
    !BRW21.ResetFromFile()
    !BRW21.ResetSort(True)

    !Now calculate the school
    SQLScript2{Prop:SQL}= 'Select PACKAGEDETAILNO, JOBCARDITEMDESCRIPTION, QTY, TOTAL from JobCardItem where JOBCARDNO in (Select JOBCARDNO from Jobcard where BEGINDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and BEGINDATETIME < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ' and DEPARTMENTNO = 6)'
    SETCLIPBOARD('Select PACKAGEDETAILNO, JOBCARDITEMDESCRIPTION, QTY, TOTAL from JobCardItem where JOBCARDNO in (Select JOBCARDNO from Jobcard where BEGINDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and BEGINDATETIME < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ' and DEPARTMENTNO = 6)')
    Loop xcv# = 1 to 1000
        Next(SQLScript2)
        
        If Not Error() then
         
        Dash:PackageNo = SQLS:Field1
        Dash:PackageDescription = SQLS:Field2
        Dash:TotalQty = SQLS:Field3
        Dash:TotalValue = SQLS:Field4
        !Access:Dashboard2016.Insert()
        END
    END                
        
        SQLScript{Prop:SQL}=  'Select Sum(TOTAL) from JobCardItem where JOBCARDNO in (Select JOBCARDNO from Jobcard where BEGINDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and BEGINDATETIME < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ' and DEPARTMENTNO = 6)'
        SETCLIPBOARD('Select Sum(TOTAL) from JobCardItem where JOBCARDNO in (Select JOBCARDNO from Jobcard where BEGINDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and BEGINDATETIME < <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>' & ' and DEPARTMENTNO = 6)')
        
    Next(SQLScript)
    If Not Error() then
        LOC:SchoolTotal = SQL:Field1
    END
           
    Display()

    !BRW21.ResetFromFile()
    !BRW21.ResetSort(True)

    !ThisWindow.Reset

    SetCursor()

LocProc.ExecuteDocumentQuery       Procedure()
    CODE

    message('Empty')        

LocProc.ExecuteShortsQuery       Procedure()
    CODE

    SetCursor(Cursor:Wait)

	SQLScript{Prop:SQL}= 'Select Count(PRINTSERVICEDETAILNO) from PrintService where CeremonyOrShort = 1 and ServiceType = 1 and ShortPrintType = 1 and CREATEDDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and CREATEDDATETIME <= <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>'
    Next(SQLScript)
    If Not Error() then
        LOC:CountFullShorts = SQL:Field1
    End

		
	SQLScript3{Prop:SQL}= 'Select Count(PRINTSERVICEDETAILNO) from PrintService where CeremonyOrShort = 1 and ServiceType = 1 and ShortPrintType = 2 and CREATEDDATETIME >= <39>' & Format(FromDate,@d10-) & ' 00:00:00.00<39>' & ' and CREATEDDATETIME <= <39>' & Format(ToDate,@d10-) & ' 23:59:59.999<39>'
	Next(SQLScript3)
		
    If Not Error() then
        LOC:CountInOrderShort = SQLSS:Field1
	END

	
	Display()
		
    SetCursor()
	
LocProc.LoadPhotosTakenDateInstitute PROCEDURE()
CODE
    setcursor(cursor:wait)

	SQLScript8{Prop:SQL}= 'call [PhotosTakeByInstituteByDate](<39>' & format(FromDate,@d10-) & '<39>,<39>' & Format(ToDate,@d10-) & '<39>,' & LOC:InstituteNo &')'
	loop
		next(SQLScript8)
        If errorCode() then BREAK.
		
		LOC:CountedPhotostaken = SQLS8:Field1
    END
    
	Display()
	
    setcursor()    

LocProc.LoadPhotosTakenDate PROCEDURE()
CODE
    setcursor(cursor:wait)

	SQLScript8{Prop:SQL}= 'call [PhotosTakeByDate](<39>' & format(FromDate,@d10-) & '<39>,<39>' & Format(ToDate,@d10-) & '<39>)'
    loop
        next(SQLScript8)
        If errorCode() then BREAK.
		
		LOC:CountedPhotostaken = SQLS8:Field1
    END
    
	Display()
	
    setcursor()    

LocProc.LoadPhotosTakenInstitute PROCEDURE()
CODE
    setcursor(cursor:wait)

    SQLScript8{Prop:SQL}= 'call [PhotosTakeByInstitute](' & LOC:InstituteNo &')'
    loop
        next(SQLScript8)
        If errorCode() then BREAK.
		
		LOC:CountedPhotostaken = SQLS8:Field1
    END
    
	Display()
	
    setcursor()    

GetConfig           procedure (string s1, string s2)
    CODE
        return 'GetConfig'
    
PutConfig           procedure (string s1, string s2)
    CODE
    
    


