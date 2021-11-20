'Name = Liberty Basic v4.5.1 Help Lab and Project Organizer v1.0
'Author(s) - , cundo, Carl Gundel,  xxgeek
'Date - Nov 2021
'_Visit the Liberty Basic forums @ https://libertybasiccom.proboards.com/ for more information.
'_Purpose - To help new users learn to code in Liberty Basic v4.5.1

'_This program is a collection of other programs written by other users,_
'_ and by Carl Gundel(creator of the Liberty Basic language)_
'_Credit goes to cundo for the lbsearch (an engine to search the lb help files) _
'_Credit also goes to cundo for the fastcode code generator that creates _
'_the shell "Window" Code _
'_Credit goes to Rod for SpriteCreator and many answers to my questions _
' and must add, a shared enthusiasm and passion for coding and helping others
'_Credit goes to Carl Gundel for his Dictionary code and his quick responses _
'_to questions regarding Liberty Basic's nuances and buried details _
'_Credit goes to tsh73 for his inspiration, advice, help, and most of all his _
'_code "proof of concept" which was the initial code to show that a TKN file _
'_can be created programatically that got this project moving forward. _
'_Credit goes to the following members who helped with all questions posed while _
'_learning to code in Liberty Basic v4.5.1 _
'_B+ for his many code samples and his desire to help others learn to code,
' not to mention his exceptional skills regarding math +(geomentry) and graphics coding
' enzo for his enthusiasm, his superb ideas, his willingness to help others, and the
' "OutoftheBox" kind of thinking he exibits
'_code (yes "code" is a member) he writes some good code, and code helps me out
' as well, so he gets a mention here too.
' Hope I didn't forget anyone, if I did, it's cause my memory can fail me.
' Sorry, let me know if I forgot YOU.

'All are members of the Liberty Basic  forums, and waiting to help YOU when you
' decide to code in Liberty Basic Just ask for help at the Liberty Basic forums and
'they (and I) will help with whatever you need regarding Liberty Basic

'xxgeek code
 'check Liberty Basic v4.5.1 Default Install Dir for existence
on error goto [errorReport]
  open "lablog.log" for append as #lablog : lablogIsOpen = 1
  #lablog, ""
   #lablog, ""
   #lablog, date$();"  ";time$()
   #lablog, ""
   #lablog, "Start Up  >>> declaring Liberty Basic install path ........."
 lbpath$ = "C:\Program Files (x86)\Liberty BASIC v4.5.1"
 #lablog, "Verifying Liberty Basic install path ........."
 res=pathExists(lbpath$)
'if Liberty Basic v4.5.1 is NOT installed to it's Default Install Dir, get Path from User using folder dialog
    if res then [start] else notice chr$(13)+" Liberty Basic v4.5.1 was not installed to the default install folder." +chr$(13)+"Hit [ok], then Select the Folder Liberty Basic v4.5.1 is Installed"

'if folder path chosen by user for Liberty Basic install is wrong catch error later with check for lbrun2.exe
#lablog, " User Liberty Basic install path   >>>  not default >>>> opening FolderDialog........."
caption$ = "Select your Liberty Basic v4.5.1 install Dir"
a$ = FolderDialog$(caption$)
if right$(FolderDialog$,1) = "\" then FolderDialog$ = left$(FolderDialog$, len(FolderDialog$)-1)
   if FolderDialog$ = "" then notice "Liberty Basic must be installed to continue" : close #lablog : end
   lbpath$ = FolderDialog$
 #lablog, "lbpath$ =  ";lbpath$ 

 [start]
 cursor hourglass
'declaring globals, and dim arrays for key$ and info$ 
    #lablog, " >>> declaring globals ........."
 global ,fixedtime$, fixeddate$, addfastfunction, selectedKey$, project, fnamenobas$, DestPath$, DestPath1$, FolderDialog$, pnum, fastfuncs$, lbexe$, lbpath$, spritecreated, lbReservedWords$, dictionary$, keyCount, q$, fileToCheck$, lastKey$, helpFilePath$, resetsearch, closehtml, categorie$, upath$, selectedpath$
  dim key$(1000)
  dim info$(500, 500)

'declare variables
#lablog, "declaring variables......."
     q$ = chr$(34) 'double quotes to wrap quotes in code
     project = 1 ' "Keep Project" defaults to true
     closehtml = 1 'close all htmlviewer windows defaults to true
     incbas = 1 'include bas file defaulkts to true
     p = 0 'passworded exe defaults to false
     lbexe$ = "liberty.exe"
    lbruntime$ = "run451.exe"
    lbReservedWords$ = "AND, APPEND, AS, BEEP, BMPBUTTON, BMPSAVE, BOOLEAN, BUTTON, BYREF, CALL, CALLBACK, CALLDLL, CALLFN, CASE, CHECKBOX, CLOSE, CLS, COLORDIALOG, COMBOBOX, CONFIRM, CURSOR, DATA, DIALOG, DIM, DLL, DO, DOUBLE, DUMP, DWORD, ELSE, END, ERROR, EXIT, FIELD, FILEDIALOG, FILES, FONTDIALOG, FOR, FUNCTION, GET, GETTRIM, GLOBAL, GOSUB, GOTO, GRAPHICBOX, GRAPHICS, GROUPBOX, IF, INPUT, INPUTCSV, KILL, LET, LINE, LISTBOX, LOADBMP, LONG, LOOP, LPRINT, MAINWIN, MAPHANDLE, MENU, NAME, NEXT, NOMAINWIN, NONE, NOTICE, ON, ONCOMERROR, OR, OPEN, OUT, OUTPUT, PASSWORD, PLAYMIDI, PLAYWAVE, POPUPMENU, PRINT, PRINTERDIALOG, PROMPT, PTR, PUT, RADIOBUTTON, RANDOM, RANDOMIZE, READ, READJOYSTICK, REDIM, REM, RESTORE, RESUME, RETURN, RUN, SCAN, SELECT, SHORT, SORT, STATICTEXT, STOP, STOPMIDI, STRUCT, SUB, TEXT, TEXTBOX, TEXTEDITOR, THEN, TIMER, TITLEBAR, TRACE, ULONG, UNLOADBMP, UNTIL, USHORT, VOID, WAIT, WINDOW, WEND, WHILE, WORD, XOR, ABS(, ACS(, AFTER$(, AFTERLAST$(, ASC(, ASN(, ATN(, CHR$(, COS(, DATE$(, DECHEX$(, EOF(, EVAL(, EVAL$(, EXP(, HBMP(, HEXDEC(, HTTPGET$(, HWND(, INP(, INPUT$(, INPUTTO$(, INSTR(, INT(, LEFT$(, LEN(, LOF(, LOG(, LOWER$(, MAX(, MIDIPOS(, MID$(, MIN(, MKDIR(, NOT(, REMCHAR$(, REPLSTR$(, RIGHT$(, RMDIR(, RND(, SIN(, SPACE$(, SQR(, STR$(, TAB(, TAN(, TIME$(, TRIM$(, TXCOUNT(, UPPER$(, UPTO$(, USING(, VAL(, WINSTRING(, WORD$(, BackgroundColor$, ComboboxColor$, CommandLine$, DefaultDir$, DisplayHeight, DisplayWidth, Drives$, Err, Err$, ForegroundColor$, Joy1x, Joy1y, Joy1z, Joy1button1, Joy1button2, Joy2x, Joy2y, Joy2z, Joy2button1, Joy2button2, ListboxColor$, Platform$, PrintCollate, PrintCopies, PrinterFont$, PrinterName$, TextboxColor$, TexteditorColor$, Version$, WindowHeight, WindowWidth, UpperLeftX, UpperLeftY"
    DllList$="vbas31w.sll vgui31w.sll voflr31w.sll vthk31w.dll vtk1631w.dll vtk3231w.dll vvm31w.dll vvmt31w.dll"
    savedProjects$ = "savedProjects"
    MyProjects$ = "MyProjects"
    programs$ = "Programs"
    vbs$ = "VB"
    ps1$ = "PS1"
    cmd$ = "CMD"
    js$ = "JS"
    html$ = "HTML"
    examples$ = "Examples"
    snippets$ = "Snippets"
    notes$ = "Notes"
    help$ = "Help"
    subroutines$ = "Subroutines"
    functions$ = "Functions"

'cundo's fastcode generator
[fastcode]
print "lb install path = ";lbpath$
#lablog,"Got lb install path using FolderDialog function";lbpath$
openhelp$ = lbpath$;"\lb4help\LibertyBASIC_4_web\amber_menu.htm"
    helpFilePath$ = lbpath$;"\lb4help\LibertyBASIC_4_web"
    helpFileMenu$ = "amber_menu.htm"
#lablog, "loading window types array......."
 dim windowTypes$(19)
    windowTypes$(0)= "":windowTypes$(1)= "dialog":windowTypes$(2)= "dialog_fs":windowTypes$(3)= "dialog_nf":windowTypes$(4)= "dialog_nf_fs"
    windowTypes$(5)= "dialog_ns_modal":windowTypes$(6)= "dialog_modal":windowTypes$(7)= "dialog_popup":windowTypes$(8)= "graphics"
    windowTypes$(9)= "graphics_fs":windowTypes$(10) = "graphics_nf":windowTypes$(11)= "graphics_nsb":windowTypes$(12)= "graphics_nsb_nf"
    windowTypes$(13)= "text":windowTypes$(14)= "text_fs":windowTypes$(15)= "text_nsb":windowTypes$(16)= "text_nsb_ins":windowTypes$(17)= "window"
    windowTypes$(18)= "window_nf":windowTypes$(19)= "window_popup"

#lablog, "beginning to create help lab window......."
nomainwin
    WindowWidth = 1368:WindowHeight = 768
    UpperLeftX= int((DisplayWidth-WindowWidth)/2)
    UpperLeftY= int((DisplayHeight-WindowHeight)/2)
    BackgroundColor$ = "lightgray"
    ForegroundColor$ = "black"

'cundo's search code - loading search array
#lablog, "- loading help list array"
  dim helpList$(500), searchList$(500)
#lablog, "helpFilePath$ = ";helpFilePath$
#lablog, "helpFileMenu$ =  ";helpFileMenu$ 
  open helpFilePath$;"\";helpFileMenu$ for input as #1
  txt$ = input$(#1, lof(#1))
  close #1
  lowerTxt$= lower$(txt$)
  #lablog, "entering search array loop  "
 while 1
  scan
    startAt = c+1
    a = instr(lowerTxt$, "href",startAt)
    b = instr(lowerTxt$, ">",a+1)
    c = instr(lowerTxt$, "</a>",b+1)
    if a=0 or b=0 or c= 0 then exit while
    hrefA= instr(lowerTxt$,chr$(34),a+1)
    hrefB= instr(lowerTxt$,chr$(34),hrefA+1)
    idx = idx +1
    helpList$(idx)= trim$(mid$(txt$,b+1,c-b-1));chr$(0);_
    trim$(mid$(txt$,hrefA+1,hrefB-hrefA-1))
      wend
  #lablog, "exiting search array loop  "

'lbsearch code by cundo
#lablog, "finishing help lab window......."
'top menu
    menu #main, "File" , "Open a File in Liberty Basic v4.5.1", [openFile], "Exit", [quit.main]
    menu #main, "Edit"
    menu #main, "View", "This Help Lab's Progress Log", [labLog], "Liberty Basic Error Log", [lberrorLog], "Runtime Error Log", [runtimeLog], "Help Lab Error Log", [helplaberrorLog]
    menu #main, "Tools" , "BAS <2> EXE", [makeEXE], "BAS <2> TKN",  [bas2tkn], ".BAS Line Count", [numofLines], "MSPaint", [pictures], "Voice Recorder", [record], "Notepad", [openNotePad]
    menu #main, "Options", "Preferences", [preferences]
    menu #main, "Browse" , "My Projects", [projectsDir],"My EXE Files", [exeDir], "My TKN Files", [tknDir], "LB Example Files", [lbexamplesDir], "LB BMP Files", [bmpDir], "LB Sprite Files", [spritesDir]
    menu #main, "Help" , "Help", [lbHelpLabHelp], "About", [about]
'lbsearch by cundo
    listbox #main.listbox1, helpList$(, lbDoubleClick, 420, 80, 320, 282
    statictext  #main.searchtext, "Search For KeyWord(s)", 755, 40, 160, 20
    statictext  #main.searchheader, "Search Liberty Basic v4.5.1 Help", 600, 7, 660, 20
    statictext  #main.searchin, "Liberty Basic Help Files", 493, 41, 185, 20
    statictext  #main.clickTip1, "Single Click to Select", 520, 67, 170, 15
    statictext  #main.clickTip2, "Single Click to Select", 845, 102, 170, 15
    textbox #main.tb, 910, 35, 155, 25
    listbox #main.listbox2, searchList$(, lbDoubleClick, 745, 115, 320, 245
    button #main.search, "&Start Searching in LB Help", buttonClick, UL, 800, 70, 220, 25
    button #main.openhelp, "Open LB &Help", [openhelp], UL, 420, 365, 95, 20
    button #main.tutorial, "Open LB T&utorial", [opentutorial], UL, 523, 365, 115, 20
    button #main.newfeatures, "New Features", [newfeatures], UL, 645, 365, 95, 20
    checkbox #main.closehtml, "Close Help (htlmviewer) Windows on Exit", [setclosehtml], [resetsearchclosehtml], 770, 365, 285, 20
'fastcode code by cundo
    texteditor  #main.ed, 8, 150, 400, 200
    statictext  #main.fastcode, "Create Window Code", 135, 5, 165, 20
    statictext  #main.st1, "< Name && Handle >", 150, 25, 128, 20
    statictext  #main.st1, "Window Type >>>", 20, 55, 115, 20
    textbox     #main.txt1, 290, 20, 115, 20
    textbox     #main.txt2, 20, 20, 115, 20
    button      #main.button1, "&Generate Code ^ + > Copy to Clipboard", dummy, ul, 70, 355, 270, 25
    combobox    #main.combo, windowTypes$(, dummy, 145, 50, 140, 20
    checkbox    #main.r1, "Use Labels instead of Subs", dummy, dummy, 8, 80, 222, 20
    button #main.freeform, "Open Liberty Basic &FreeForm Editor", [openFreeForm], UL, 90, 120, 250, 25
'dictionary code by Carl Gundel
    texteditor #main.value, 293, 415, 770, 230
    listbox #main.keys, keys$(), [keySelected], 5, 415, 285, 205
'xxgeeks code
'category radio buttons
   radiobutton #main.programs, programs$, [progs], resetHandler, 400, 395, 80, 20
   radiobutton #main.examples, examples$, [exams], resetHandler, 500, 395, 80, 20
   radiobutton #main.snippets, snippets$, [snipps], resetHandler, 600, 395, 70, 20
   radiobutton #main.savedprojects, MyProjects$, [projs], resetHandler, 290, 395, 90, 20
   radiobutton #main.notes, notes$, [notes], resetHandler, 905, 395, 55, 20
   radiobutton #main.subroutines, subroutines$, [subroutines], resetHandler, 695, 395, 100, 20
   radiobutton #main.functions, functions$, [functions], resetHandler, 805, 395, 80, 20
   radiobutton #main.VBS, vbs$, [vbs], resetHandler, 1070, 445, 45, 20
   radiobutton #main.PS1, ps1$, [ps1], resetHandler, 1070, 480, 45, 20
   radiobutton #main.CMD, cmd$, [cmd], resetHandler, 1070, 515, 45, 20
   radiobutton #main.JS, js$, [js], resetHandler, 1070, 550, 45, 20
   radiobutton #main.html, html$, [html], resetHandler, 1070, 585, 60, 20
   radiobutton #main.help, help$, [help], resetHandler, 1070, 625, 60, 20
'buttons bottom left and middle
   button #main.addListing, "&New ";left$(categorie$, (len(categorie$) - 1)), [newKey], ul, 5, 624, 137, 23
   button #main.deleteListing, " &Delete Selected ", [deleteKey], ul, 150, 624, 140, 23
   button #main.makeproject, " &Make New Project", [makeproject], ul, 5, 624, 140, 23
   button #main.remakeproject, "Re&-Make Selected", [remakeproject], ul,5, 652, 140, 23
   button #main.runListing, "&Run Selected", [runKey], ul, 150, 652, 140, 23
   button #main.runlb, "Edit Selected in Libert&y Basic", [edit_In_lb_IDE], ul, 330, 652, 200, 23
   button #main.editInNotepad, "Edit Se&lected with Notepad", [editInNotepad], ul, 580, 652, 200, 23
'create buttons
   button #main.lbProgs, "&Liberty Basic", [lbProgs], ul, 1085, 30, 120, 25
   button #main.freeform, " LB &FreeForms ", [openFreeForm], ul, 1220, 30, 120, 25
   button #main.pictures, "MS&Paint", [pictures], ul, 1085, 65, 120, 23
   button #main.sprites, " Spr&iteCreator ", [sprites], ul, 1220, 65, 120, 23
   'useful tools buttons
   button #main.task, " Task&_Manager ", [taskman], ul, 1085, 125, 120, 23
   button #main.calc, " &Calculator ", [calc], ul, 1085, 155, 120, 23
   button #main.record, " &Voice Recorder ", [record], ul, 1085, 185, 120, 23
   button #main.basexe, "  &BAS < 2 > EXE ", [makeEXE], ul, 1085, 215, 120, 23
   button #main.bas2tkn, " BAS < 2 > &TKN ", [bas2tkn], ul, 1085, 245, 120, 23
   button #main.numLines, " .BAS Line Count ", [numofLines], ul, 1085, 275, 120, 23
   button #main.notepad, " NoteP&ad ", [openNotePad], ul, 1085, 305, 120, 23
   button #main.lberrorlog, "LB Error Log", [lberrorLog], ul, 1085, 340, 85, 23
   button #main.helplablog, "Lab Log", [labLog], ul, 1280, 340, 60, 23
   button #main.runtimelog, "Runtime Log", [runtimeLog], ul, 1177, 340, 95, 23
'browse buttons
   button #main.projectsDir, " My Projects ", [projectsDir], ul, 1220, 125, 120, 23
   button #main.exeDir, " My EXE Files ", [exeDir], ul, 1220, 155, 120, 23
   button #main.tknDir, " My TKN Files ", [tknDir], ul, 1220, 185, 120, 23
   button #main.bmpDir, " BMP Folder ", [bmpDir], ul, 1220, 215, 120, 23
   button #main.spritesDir, " Sprites Folder ", [spritesDir], ul, 1220, 245, 120, 23
   button #main.defaultDir, " DefaultDir$ ", [defaultDir], ul, 1220, 305, 120, 23
   button #main.examplesDir, " Examples Folder ", [lbexamplesDir], ul, 1220, 275, 120, 23
'ascii chart combo box and static text
   statictext  #main.asciitext, "ASCII Codes chr$()", 1220, 375, 125, 15
   combobox #main.asciiList, asciiList$(), asciiSelected , 1215, 390, 125, 40
'lb samples  to reserved words comboboxes and statictext
   statictext  #main.lbsamplestext, "LB Examples", 1140, 420, 125, 15
   combobox #main.lbsamplesList, lbsamplesList$(), lbsampleSelected , 1120, 435, 120, 25
   statictext  #main.lbdialogstext, "LB Dialogs", 1255, 420, 100, 15
   combobox #main.lbdialogsList, lbdialogsList$(), lbdialogSelected , 1245, 435, 100, 25
   statictext  #main.reservedwordstext, "Reserved Words", 1130, 465, 120, 15
   combobox #main.lbreservedwordsList, lbreservedwordsList$(), lbreservedwordSelected , 1120, 480, 125, 25
'lb bak(up) Files combobox and statictext
   statictext  #main.lbbaktext, "LB BAK (UP) Files", 1095, 375, 120, 15
   combobox #main.lbbakfilesList, lbbakfilesList$(), lbbakfileSelected , 1090, 390, 120, 25
'right side static text
   statictext  #main.creators, "Create With", 1160, 5, 160, 20
   statictext  #main.useful, "Useful Tools", 1095, 102, 160, 20
   statictext  #main.browse, "Browse", 1250, 100, 162, 20
   statictext  #main.choose, "Select  a Category >>  >>> ", 55, 395, 200, 20
   statictext  #main.killtext, "Kill All LB Processes >", 1165, 665, 150, 20
   statictext  #main.logsClear, "Clear all Logs >", 1205, 628, 110, 20
   statictext  #main.lbforums, "Visit https://libertybasiccom.proboards.com/", 05, 700, 275, 25
'kill all button
   button #main.killAll, " &K ", [killAll], UL, 1315, 658, 30, 30
   button #main.clearLogs, " &X ", [clearLogs], UL, 1315, 620, 30, 30

#lablog, "opening help lab window......."
 open "Liberty Basic v4.5.1 Help Lab and Project Organizer" for window as #main

#lablog, "defining fonts, visibility, other attributes in help lab window......."
 #main "trapclose [quit.main]"
        #main.addListing, "!hide"
        #main.remakeproject, "!hide"
        #main.makeproject, "!hide"
        #main.deleteListing, "!hide"
        #main.runListing, "!hide"
        #main.runlb, "!hide"
        #main.editInNotepad, "!hide"
        #main.listbox1 "singleclickselect"
        #main.listbox2 "singleclickselect"
        #main "font arial 10 Bold"
        #main.asciiList, "font arial 10 bold"
        #main.txt1 "#main"
        #main.txt2 "untiltled"
        #main.r1 "set"
        #main.closehtml, "set"
        #main.listbox1, "font arial 12 bold"
        #main.listbox2, "font arial 12 bold"
        #main.value, "!font arial 12 bold"
        #main.keys, "font arial 12 bold"
        #main.lbProgs, "!font arial 12 bold"
        #main.freeform, "!font arial 12 bold"
        #main.ed "!font arial 12 bold"
        #main.clickTip1, "!font arial 8 bold"
        #main.clickTip2, "!font arial 8 bold"
        #main.combo "selectindex 17"
        #main.keys "singleclickselect"
        #main.button1, "!font arial 10 bold"
        #main.searchin, "!font arial 12 bold"
        #main.st1, "!font arial 10 bold "
        #main.fastcode, "!font arial 12 bold"
        #main.search, "!font arial 12 bold"
        #main.searchheader, "!font arial 12 bold"
        #main.creators "!font arial 12 bold"
        #main.useful, "!font arial 12 bold"
        #main.browse, "!font arial 12 bold"
        #main.choose, "!font arial bold 12"
        #main.sprites, "!font arial 12 Bold"
        #main.lbforums, "!font arial 8 bold"
        #main.pictures, "!font arial 12 bold"
        #main.runListing, "!hide"
        #main.deleteListing, "!hide"
        #main.remakeproject, "!font arial 10 bold"
        #main.makeproject, "!font arial 10 bold"

#lablog, "calling progressbar 0......."
  call progressBar
'get users home dir path
#lablog,"calling getUserpath......."
  call getUserPath
      pnum = 1
      #lablog, "User Home dir Path = ";upath$
      print "User Home dir Path = ";upath$
'load up the list and combo boxes for dialogs, lbsamples, lb bak files, reserved words, and ascii codes
     #lablog, "calling progressbar 1......."
     call progressBar
#lablog, "calling getAscii"
    call getAscii
     pnum = 2
#lablog, "calling progressbar 2......."
     call progressBar
#lablog, "getlbreservedwords......."
     call getlbreservedwords
     pnum = 3
#lablog, "calling progressbar 3......."
     call progressBar
#lablog, "calling getlbsamples......."
     call getlbsamples
     pnum = 4
#lablog, "calling progressbar 4......."
     call progressBar
#lablog, "calling getlbBakFiles......."
  call getlbBakFiles
#lablog, "calling progressbar 5......."
     pnum = 5
#lablog,"calling getlbdialogs......."
     call getlbdialogs
     call progressBar
     #main.tb, "!setfocus"
     cursor normal
     #lablog, "Start Up completed, waiting for user input"
  wait

'Dictionary by Carl Gundel + edited by xxgeek
[newKey]   'ask the user for a new listing
call saveValue
newKey$ = ""
print "Creating new ";categorie$;" Listing"
#lablog,"@ [newKey] - creating new ";categorie$;" Listing"
      if len(left$(categorie$, (len(categorie$) - 1))) < 4 then [notPlural]
  prompt "Enter a Name (or Title) for the New " + left$(categorie$,(len(categorie$)-1)); newKey$
   if newKey$ <> "" then [continue] else wait

[notPlural]
   prompt "Enter a Name (or Title) for the New  "+categorie$+" Script"; newKey$
      if newKey$ = "" then wait

[continue]
#lablog, "@ continue creating new ..";categorie$;" Listing >";newKey$
   if newKey$ <> "" then
        call setValueByName newKey$, ""
        call loadKeys
        #main.keys "select "; newKey$
        #main.value "!cls";
        #main.value "!setfocus";
        call collectGarbage
        call writeDictionary
        lastKey$ = newKey$
'xxgeek code
        if tkn = 2 or tkn = 4 then
                #lablog "Adding new key -> ;";fname$;" to -> ";categorie$
                #main.savedprojects, "set"
                 open fname$ for input as #1
                  open categorie$ for append as #2
                  print #2, input$(#1, lof(#1));
                  close #1
                  close #2
                 call saveValue
                 call readDictionary
                call loadKeys
               call cleanup
             tkn = 0
             cursor normal
             goto [done]
    end if

        if tkn = 3 then
                #lablog "Adding new key -> ;";fname$;" to -> ";categorie$
                 #main.programs, "set"
                  open fname$ for input as #1
                   open categorie$ for append as #2
                    print #2, input$(#1, lof(#1));
                   close #1
                   close #2
                   call saveValue
                  call readDictionary
                 call loadKeys
               call cleanup
             tkn = 0
             cursor normal
           goto [done]
    end if
  end if
 wait

'xxgeek code
[openhelp]
#lablog, "opening selected help file with htmlviewer (if available) or default browser (if not)"
if fileExists(DefaultDir$, "htmlviewer.exe") <> 0 then
 run "htmlviewer.exe ";openhelp$
       else
        run "explorer.exe ";openhelp$
   end if
 wait

[opentutorial]
#lablog, "opening the Liberty Basic Tutorial "
lb4tutPath$ = "Application Data\Liberty BASIC v4.5.1"
lb4tut$ = "lb4tutorial.lsn"
   if fileExists(upath$;"\";lb4tutPath$, lb4tut$ ) <> 0 then
   runFile$ = upath$;"\";lb4tutPath$;"\";lb4tut$ 
          run lbpath$;"\";lbexe$;" ";runFile$
          else
          notice "Can't find Liberty Basic Tutorial"
   end if
 wait

[newfeatures]
#lablog, "opening Liberty Basic NewFeatures"
lb4tutPath$ = "Application Data\Liberty BASIC v4.5.1"
lb4tut$ = "new4features.lsn"
   if fileExists(upath$;"\";lb4tutPath$, lb4tut$ ) <> 0 then
   runFile$ = upath$;"\";lb4tutPath$;"\";lb4tut$ 
          run lbpath$;"\";lbexe$;" ";runFile$
          else
          notice "Can't find Liberty Basic New Features"
   end if
 wait

[openFreeForm]
#lablog, "opening the LB FreeForm editor"
    freeForm$ = upath$;"\Application Data\Liberty Basic v4.5.1\freeform450.bas"
  run lbpath$;"\";lbexe$;" -R -A ";freeForm$
 wait

[killAll]
   answer$ = "YES"
   prompt "Shuting down ALL Liberty.exe Files "+chr$(13)+"Is Your Work Saved - ARE YOU SURE?";answer$
  if answer$ <> "YES" then wait
 #lablog, "@ - [killAll] - calling sub saveValue"
  call saveValue
 #lablog, "@ - [killAll] - calling sub cleanup"
  call cleanup
 #lablog, "@ - [killAll] - killing all open htmlviewer tasks"
  if fileExists(DefaultDir$, "htmlviewer.exe") <> 0 then run "taskkill /IM htmlviewer.exe /F", HIDE
 #lablog, "@ - [killAll] killing all open liberty basic tasks"
  run "taskkill /IM liberty.exe /F", HIDE
 end

'Carl Gundel's Dictionary code
[keySelected] 'a key in the list was selected
#lablog, "@ - [keySelected] - a key was selected "
  call saveValue
    #main.value "!origin 0, 0"
    #main.keys "selection? selectedKey$"
#lablog, "selection made was  ";selectedKey$ 
        print "selectedKey$ = ";selectedKey$
      selectedValue$ = getValue$(selectedKey$)
  #main.value "!contents selectedValue$";
     lastKey$ = selectedKey$
     #main.value "!setfocus";
     #main.value, "!origin 0 0"
 wait

'xxgeek code
[deleteKey]
print "@- [deleteKey]....."
    #main.keys "selection? selectedKey$"
    #lablog, "selected key was  ";selectedKey$
#lablog, "@- [deleteKey] deleting ";selectedKey$;" from ";categorie$
 if pathExists(DefaultDir$;"\BAK") = 0 then res = mkdir(DefaultDir$;"\BAK")
      categor$ = categorie$
  categor$ = left$(categor$, len(categor$)-1)
   answer$ = "yes"
    if selectedKey$ = "" then
    prompt "No Selection made. Try again?";answer$
    if selectedKey$ <> "yes" then wait
    end if
  prompt "Deleting Entry" + chr$(13) + selectedKey$ + "   OK ?";answer$
    if answer$ <> "yes" then wait
 cursor hourglass
    #main.value, "!selectall"
    #main.value, "!cut"
    deleteIt$ = selectedKey$
  call saveValue
    #lablog, "copying ";categorie$;" to ";"BAK Dir appending date and time......."
   open categorie$ for input as #1
    open "BAK\";categorie$;fixeddate$;fixedtime$ for output as #2
   #2, input$(#1, lof(#1));
  close #2
  close #1
   #lablog, "Opening tempfile to find line to delete from ";categorie$
   open categorie$ for input as #1
    tempfile$ = "tempfile"
   open tempfile$ for output as #2
   #lablog, "tempfileOpened =";tempfile$
    word1$ = chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134) + selectedKey$ + chr$(134);chr$(165);chr$(134);"value";chr$(134);chr$(165);chr$(134)
 #lablog, "opening window for delay while deleting ";selectedKey$;" from ";categorie$
WindowWidth = 600:WindowHeight = 150
    UpperLeftX= int((DisplayWidth-WindowWidth)/2)
    UpperLeftY= int((DisplayHeight-WindowHeight)/2)+185
    BackgroundColor$ = "lightgray"
    ForegroundColor$ = "black"
statictext #deleteKey.text, " Please wait - Deleting a listing can take some time, depending on the List and the size of each Listing's file. Approximatlely 1 minute for each 160 KB of data", 10, 40, 580, 100
      open "Information Message" for dialog_popup                                                                                 as #deleteKey
 #deleteKey "trapclose [quit.deleteKey]"
 #deleteKey.text "!font arial 12 bold"
 while eof(#1) = 0
      line input #1, line$
      print "dontsave = ";line$
      print "word1$ = ";word1$
      print "line$ = ";line$
      if line$ = word1$ then [dontSave]
      #2, line$
[dontSave]
#lablog, "deleting line ";line$
 wend
     close #1
     close #2
 #lablog, "deleting ";selectedKey$
     deleteMe$ = categorie$;".bak"
   if fileExists(DefaultDir$, deleteMe$) = 0 then
      print "attempting to name categorie$ as categorie$.bak"
       name categorie$ as categorie$;".bak"
        print "Finished renaming";categorie$;" To  ";categorie$;".bak"
         name tempfile$ as categorie$ 
   else
         kill deleteMe$
         print "deleted old bak file"
         name categorie$ as categorie$;".bak"
        name tempfile$ as categorie$
       print "renamed temp file as";categorie$
  end if
     selectedKey$ = ""
     lastKey$ = ""
     call readDictionary
     call collectGarbage
     call loadKeys
     call saveValue
#lablog, "finished deleting Listing ";selectedKey$;" in ";categorie$
[quit.deleteKey]
close #deleteKey
cursor normal
wait

'run selected MyProjects, or Program
[runKey]
if selectedKey$ = "" then notice "Select from list, try again" : wait
#lablog, "Running ";selectedKey$;" in ";categorie$;" List"
#lablog, "selected Key = ";selectedKey$
      print "selected Key = ";selectedKey$
      runFile$ = savedProjects$;"\";selectedKey$;"\";selectedKey$;".exe"
 if fileExists(DefaultDir$;"\";savedProjects$;"\";selectedKey$, selectedKey$;".exe") <> 0 then
 #lablog, "running selectedKey =  ";savedProjects$;"\";selectedKey$;"\";selectedKey$;".exe"
     run q$;DefaultDir$;"\";savedProjects$;"\";selectedKey$;"\";selectedKey$;".exe";q$
  else
     Prompt "Project NOT Created.....yet?" + chr$(13) + "Make Project and Try again" ; answer$
  end if
wait

'open selected listing in Liberty Basic IDE
[edit_In_lb_IDE]
if selectedKey$ = "" then notice "Select from list, try again" : wait
#lablog, "Running ";selectedKey$;" of ";categorie$;" with lb IDE"
   print "selectedKey$ = ";selectedKey$
      runFile$ = DefaultDir$;"\";savedProjects$;"\";selectedKey$;"\";selectedKey$;".bas"
         print "runFile$ = ";runFile$
         res = fileExists(DefaultDir$;"\";savedProjects$;"\";selectedKey$, selectedKey$;".bas")
    if res then
         run lbpath$;"\";lbexe$;" ";q$;runFile$;q$
   else
         #main.value, "!contents? valueNow$";
        open "untitled.bas" for output as #1
          #1, valueNow$
        close #1
        tempfile$ = DefaultDir$;"\untitled.bas"
       print lbpath$;"\";lbexe$;" ";q$;tempfile$;q$;" = ";lbpath$;"\";lbexe$;" ";q$;tempfile$;q$
      run lbpath$;"\";lbexe$;" ";q$;tempfile$;q$
   end if
wait

[editInNotepad]
#lablog, "Running ";selectedKey$;" of ";categorie$;" with Notepad"
if selectedKey$ = "" then notice "Select from list, try again" : wait
 runFile$ = DefaultDir$;"\";savedProjects$;"\";selectedKey$;"\";selectedKey$;".bas"
         print "runFile$ = ";runFile$
         res = fileExists(DefaultDir$;"\";savedProjects$;"\";selectedKey$, selectedKey$;".bas")
    if res then
         run "notepad ";q$;runFile$;q$
         #lablog, "Notepad opened sucessfully ";selectedKey$;" of ";categorie$;" FILE USED, not TEXT"
   else
   answer$ = "yes"
   prompt "No project made yet for "+ chr$(13)+ selectedKey$ + chr$(13) + "Edit the text for " + chr$(13) + selectedKey$ +"in notepad ?" ; answer$
if answer$ <> "yes" then wait
        #main.value, "!contents? forNotepad$";
        open "untitledTemp.bas" for output as #1
        #1, forNotepad$
       close #1
      run "notepad ";q$;"untitledTemp.bas";q$
#lablog, "Notepad opened sucessfully ";selectedKey$;" of ";categorie$;" TEXT used, not FILE"
         end if
wait

'top menu "Open File in liberty basic  IDE"
[openFile]
#lablog "@- [openFile] opening filedialog to select a file for edit in LB IDE"
 filedialog "Open \ Select a Liberty Basic Source File (.bas) ", DefaultDir$; "\*.bas", openFilename$
    if openFilename$ = "" then wait
  run lbpath$;"\";lbexe$;" ";openFilename$
 wait

'set "close htmlviewer" checkbox
[setclosehtml] closehtml = 1 : print "closehtml [setclose] =  ";closehtml : wait

'reset "close htmlviewer" checkbox
[resetsearchclosehtml] closehtml = 0 : print "closehtml  [resetsearchclose] =  ";closehtml : wait

[runtimeLog]
#lablog, "@ [runtimeLog] - opening notepad to view error log......"
if fileExists(DefaultDir$, "error.log") <> 0 then
 run "notepad ";q$;DefaultDir$;"\";"error.log";q$
else
answer$ = "YES"
prompt "No Runtime errors have occured yet"+chr$(13)+ "Was this Helpful?";answer$
end if
wait

'radio button selections from MyProjects to Help
[projs]
#lablog "@- [projs] ........"
   #main.runListing, "!show"
   #main.makeproject, "!show"
   #main.runlb, "!show"
   #main.addListing, "!hide"
   #main.editInNotepad, "!show"
   #main.deleteListing, "!show"
   call saveValue
   #main.value, "!cls"
      categorie$ = MyProjects$
   #main.remakeproject, "Re-Make Project"
   #main.remakeproject, "!show"
 call resetRadioOptions
    category$ = categorie$
    category$= left$(category$, (len(category$) - 1))
    category$ = right$(category$,7)
   #main.addListing, "&New ";category$
   #main.choose, "Select a  ";category$
    wait

[progs]
#lablog "@- [progs] ............"
   #main.runListing, "!show"
   #main.makeproject, "!hide"
   #main.runlb, "!show"
   #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.editInNotepad, "!show"
  call saveValue
   #main.value, "!cls"
     categorie$ = programs$
     #main.remakeproject, "Re-Make Program"
        #main.remakeproject, "!show"
  call resetRadioOptions
   #main.addListing, "&New ";left$(categorie$, (len(categorie$) - 1))
     category$ = categorie$
     category$= left$(category$, (len(category$) - 1))
   #main.choose, "Select a  ";category$
   wait

[exams]
#lablog "@- [exams] ............"
   #main.runListing, "!hide"
   #main.runlb, "!show"
   #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.editInNotepad, "!show"
  call saveValue
    #main.value, "!cls"
    categorie$ = examples$
  call resetRadioOptions
  #main.addListing, "&New ";left$(categorie$, (len(categorie$) - 1))
    category$ = categorie$
 category$= left$(category$, (len(category$) - 1))
  #main.choose, "Select an  ";category$
 wait

[snipps]
#lablog "@- [snipps] ............"
   #main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.runlb, "!show"
   #main.runListing, "!hide"
   #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.editInNotepad, "!show"
 call saveValue
  #main.value, "!cls"
  categorie$ = snippets$
  call resetRadioOptions
  #main.addListing, "&New ";left$(categorie$, (len(categorie$) - 1))
 category$ = categorie$
 category$= left$(category$, (len(category$) - 1))
  #main.choose, "Select a  ";category$
 wait

[subroutines]
#lablog "@- [subroutines] ............"
   #main.runListing, "!hide"
   #main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.runlb, "!show"
   #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.editInNotepad, "!show"
   #main.editInNotepad, "!show"
  call saveValue
   #main.value, "!cls"
    categorie$ = subroutines$
  call resetRadioOptions
    category$ = categorie$
    category$= left$(category$, (len(category$) - 1))
  #main.addListing, "&New ";category$
  #main.choose, "Select a  ";category$
wait

[functions]
#lablog "@- [functions] ............"
   #main.runListing, "!hide"
   #main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.runlb, "!show"
   #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.editInNotepad, "!show"
   #main.editInNotepad, "!show"
  call saveValue
    #main.value, "!cls"
       categorie$ = functions$
       call resetRadioOptions
    #main.addListing, "&New ";left$(categorie$, (len(categorie$) - 1))
    category$ = categorie$
 category$= left$(category$, (len(category$) - 1))
  #main.choose, "Select a  ";category$
wait

[notes]
#lablog "@- notes] ............"
   #main.runListing, "!hide"
   #main.runlb, "!show"
   #main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.addListing, "!show"
   #main.editInNotepad, "!show"
   #main.deleteListing, "!show"
   #main.editInNotepad, "!show"
call saveValue
  #main.value, "!cls"
  categorie$ = notes$
 call resetRadioOptions
 #main.addListing, "&New ";left$(categorie$, (len(categorie$) - 1))
 category$ = categorie$
 category$= left$(category$, (len(category$) - 1))
  #main.choose, "Select a  ";category$
 wait

[vbs]
#lablog "@- [vbs] ............"
   #main.runListing, "!hide"
   #main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.runlb, "!show"
   #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.editInNotepad, "!show"
   #main.editInNotepad, "!show"
  call saveValue
   #main.value, "!cls"
    categorie$ = vbs$
    call resetRadioOptions
 #main.addListing, "&New ";categorie$;" Script"
 category$ = categorie$
   #main.choose, "Select a  ";category$;"  Script"
wait

[ps1]
#lablog "@- [ps1] ............"
   #main.runListing, "!hide"
   #main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.runlb, "!show"
   #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.editInNotepad, "!show"
   #main.editInNotepad, "!show"
  call saveValue
   #main.value, "!cls"
  categorie$ = ps1$
  call resetRadioOptions
  #main.addListing, "&New ";categorie$;" Script"
   category$ = categorie$
   #main.choose, "Select a  ";category$;"  Script"
wait

[cmd]
#lablog "@- [cmd] ............"
   #main.runListing, "!hide"
#main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.runlb, "!show"
   #main.editInNotepad, "!show"
  call saveValue
   #main.value, "!cls"
  categorie$ = cmd$
  call resetRadioOptions
  #main.addListing, "&New ";categorie$;" Script"
   category$ = categorie$
   #main.choose, "Select a  ";category$;"  Script"
wait

[js]
#lablog "@- [js] ............"
   #main.runListing, "!hide"
   #main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.runlb, "!show"
   #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.editInNotepad, "!show"
  call saveValue
   #main.value, "!cls"
  categorie$ = js$
  call resetRadioOptions
  #main.addListing, "&New ";categorie$;" Script"
   category$ = categorie$
   #main.choose, "Select a  ";category$;"  Script"
wait

[html]
#lablog "@- [html] ............"
   #main.runListing, "!hide"
   #main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.runlb, "!show"
      #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.editInNotepad, "!show"
  call saveValue
   #main.value, "!cls"
  categorie$ = html$
  call resetRadioOptions
  #main.addListing, "&New ";categorie$;" Script"
   category$ = categorie$
   #main.choose, "Select a  ";category$;"  Script"
wait

[help]
#lablog "@- [help] ............"
   #main.runListing, "!hide"
   #main.makeproject, "!hide"
   #main.remakeproject, "!hide"
   #main.runlb, "!show"
   #main.addListing, "!show"
   #main.deleteListing, "!show"
   #main.editInNotepad, "!show"
call saveValue
   #main.value, "!cls"
       categorie$ = help$
       call resetRadioOptions
       category$ = categorie$
#main.addListing, "&New ";category$
   #main.choose, "Select a  ";category$;" Topic"
   wait

'open windows taskmanager (used to kill "non responsive" code (usually caught in loops)
[taskman]
#lablog "@- [taskman] ............"
run "explorer c:\Windows\System32\taskmgr.exe"
wait

'open Windows Calculator
[calc]
#lablog "@- [calc] ............"
run  "calc.exe"
wait

'open Windows Notepad
[openNotePad]
#lablog "@- [notepad] ............"
run "notepad.exe"
wait

'open Windows Voice recorder
[record]
#lablog "@- [record] ............"
run "explorer.exe shell:appsFolder\Microsoft.WindowsSoundRecorder_8wekyb3d8bbwe!App"
wait

'open the following in Windows Explorer
[projectsDir]
#lablog "@- [projectsDir] ............"
run "explorer.exe ";q$;DefaultDir$;"\";"savedProjects";q$
wait

[exeDir]
#lablog "@- [exeDir] ............"
run "explorer.exe ";q$;DefaultDir$;"\";"EXE";q$
wait

[spritesDir]
#lablog "@- [spritesDir] ............"
     if pathExists(upath$+"\AppData\Roaming\Liberty Basic v4.5.1\SPRITES") <> 0 then
         lbusppath$ = upath$+"\AppData\Roaming\Liberty Basic v4.5.1\SPRITES"
         run "explorer.exe ";q$;lbusppath$;q$
     else
           if pathExists(upath$+"\Application Data\Liberty Basic v4.5.1\SPRITES") <> 0 then
              lbusppath$ = upath$+"\Application Data\Liberty Basic v4.5.1\SPRITES"
               run "explorer.exe ";q$;lbusppath$;q$
           end if
     end if
 wait

[tknDir]
#lablog "@- [tknDir] ............"
   a$ = DefaultDir$;"\TKN"
   run "explorer.exe ";q$;a$;q$
 wait

[bmpDir]
#lablog "@- [bmpDir] ............"
   if pathExists(upath$+"\AppData\Roaming\Liberty Basic v4.5.1\bmp") <> 0 then
      lbusppath$ = upath$+"\AppData\Roaming\Liberty Basic v4.5.1\bmp"
      run "explorer.exe ";q$;lbusppath$;q$
   else
      if pathExists(upath$+"\Application Data\Liberty Basic v4.5.1\bmp") <> 0 then
        lbusppath$ = upath$+"\Application Data\Liberty Basic v4.5.1\bmp"
        run "explorer.exe ";q$;lbusppath$;q$
     end if
   end if
 wait

[lbexamplesDir]
#lablog "@- [lbexamplesDir] ............"
    if pathExists(upath$+"\AppData\Roaming\Liberty Basic v4.5.1") <> 0 then
      lbusppath$ = upath$+"\AppData\Roaming\Liberty Basic v4.5.1"
      run "explorer.exe ";q$;lbusppath$;q$
    else
       if pathExists(upath$+"\Application Data\Liberty Basic v4.5.1") <> 0 then
          lbusppath$ = upath$+"\Application Data\Liberty Basic v4.5.1"
          run "explorer.exe ";q$;lbusppath$;q$
      end if
    end if
 wait

[defaultDir]
#lablog, "Opening File Manager to : ";DefaultDir$
  run "explorer.exe ";q$;DefaultDir$;q$
 wait

[about]
#lablog "@- [about] displaying about page............"
  message$ = chr$(13);"     LB Help Lab and Project Manager v1.0";_
  chr$(13);_
  chr$(13);_
  "     Created by xxgeek";_
  chr$(13);_
  chr$(13);_
  "     Date - Oct 28 2021";_
  chr$(13);_
  chr$(13);_
  "     Purpose - To Help New Liberty Basic Coders with Information, Help,  plus Examples";_
  "     and the ability to automatically manage and organize their coded, compiled, TKN,d and EXE,d projects.";_
  "     With Functions, Subroutines, code samples, code generators, searchable Help Files,";_
  "     ASCII codes, Reserved Words, MSpaint, notepad, error logs etc, at their FingerTips";_
  chr$(13);_
  chr$(13);_
  "    Liberty Basic Help Lab and Project Manager  Is a collection of programs created by";_
  "     the creator of Liberty Basic, Carl Gundel";_
  "     and members of the Liberty Basic forums @ https://libertybasiccom.proboards.com/";_
  "     Stitched together with added programs, features and abilities to enhance and";_
  "     make more efficient, the Liberty Basic coding experience with help at the users finger tips";_
  "     by utuizing the built in abilities, and files of Liberty Basic 4.5.1 and Windows 10";_
  chr$(13);_
  chr$(13);_
  "    Credit goes to cundo for lbsearch(the LB help file search engine)";_
  "    Credit also goes to cundo for the fastcode(window code generator)";_
  "    Credit goes to Carl Gundel for the Dictionary code(handling the categories";_
  "    lists, and texteditor data saving and retrieval (not to mention Carl Gundel";_
  "    gives Liberty Basic away FREE! with runtime files so your projects are royalty FREE)";_
  "    Credit goes to Rod, not only for his SpriteCreator program, but for his many";_
  "    answers to my difficult questions, while at times going out of his way to help.";_
  "    Last, but not least, rather most importantly credit goes to a member - handle name";_
  "    tsh73 for his proof of concept code demonstrating that the TKN file can be created;";_
  "    using lb code (that got this project off the ground), his help whenever needed,";_
  "    his ideas, programming skills and his abilty to dig into posted code and fix problem issues"
 a$ = GetMessage$(message$)
wait

[lberrorLog]
 lberrorlog$ = "error.log"
#lablog "@- [lberrorLog]opening the lb error log using notepad............"
    if fileExists(upath$;"\AppData\Roaming\Liberty Basic v4.5.1", lberrorlog$) then
run "notepad ";q$;upath$;"\AppData\Roaming\Liberty Basic v4.5.1\error.log";q$
else
      if fileExists(upath$;"\Application Data\Liberty Basic v4.5.1", lberrorlog$) then
         run "notepad ";q$;upath$;"\Application Data\Liberty Basic v4.5.1\error.log";q$
      else
         notice "Can't find the LB error log"
      end if
   end if
wait

[helplaberrorLog]
#lablog "@- [helplaberrorLog] - opening helplaberrorLog using notepad........"
 helplaberrorLog$ = "lbHelpLabError.log"
  run "notepad ";helplaberrorLog$
 wait

[labLog]
#lablog "@- [labLog] User clicked to open error log - closing error log temporarily............"
   close #lablog
    lablog$ = "lablog.log"
   run "notepad ";lablog$
   open lablog$ for append as #lablog
 wait

[lbHelpLabHelp]
 help$ = "help.txt"
#lablog "@- [lbHelpLabHelp] opening notepad to.....  ";help$
   if fileExists(DefaultDir$, help$) then
       run "notepad ";help$
   else
     notice "Can't find Help File (help.txt) in DefaultDir$"
   end if
 wait

'clear all logs pressed
[clearLogs]
print "@ [clearLogs] - closing #lablog temporarily, and clearing all logs by deletion"
if lablogIsOpen = 1 then close #lablog
lablog$ = "lablog.log"
runtimeErrorLog$ = "error.log"
lbHelpLabErrorLog$ = "lbHelpLabError.log"
if fileExists(DefaultDir$, lablog$) <> 0 then kill DefaultDir$;"\";lablog$
if fileExists(DefaultDir$, lbHelpLabErrorLog$) <> 0 then kill DefaultDir$;"\";lbHelpLabErrorLog$
if fileExists(DefaultDir$, runtimeErrorLog$) <> 0 then kill DefaultDir$;"\";runtimeErrorLog$
open lablog$ for append as #lablog
#lablog, "@ [clearLogs] -  #lablog closed temporarily while clearing all logs by deletion"
#lablog, "lablog re-opened > Logs Cleared"
print "Logs Cleared"
wait

'open Liberty Basic IDE
[lbProgs]
#lablog "@- [lbProgs] opening Liberty Basic............"
  run lbpath$;"\";lbexe$
wait

'open mspaint for creating pictures (bmp, jpg, icons, etc)
[pictures]
#lablog "@- [pictures] opening mspaint............"
  run "mspaint.exe"
wait

'Rod's SpriteCreator
[sprites]
#lablog "@- [sprites] looking for spritecreator............"
 spriteEXEpath$ = DefaultDir$;"\SpriteCreator v2"
   if fileExists(spriteEXEpath$,"SpriteCreator.exe") <> 0 then
         run q$;spriteEXEpath$;"\";"SpriteCreator.exe";q$
          wait
   end if
   if fileExists(spriteEXEpath$,"SpriteCreator.bas") = 0 then
           notice chr$(13);"Download Rod's SpriteCreator at "+chr$(13)+" https://gamebin.webs.com/SpriteCreator%20v2.zip"+chr$(13)+" and extract to DefaultDir$, then Try Again" : wait
   end if
     if fileExists(spriteEXEpath$,"SpriteCreator.bas") <>  0 then [alreadyDownloaded]
wait

[alreadyDownloaded]
#lablog "@- [alreadyDownloaded] ............"
  spriteCreatorPath$ = DefaultDir$;"\SpriteCreator v2"
  res = fileExists(DefaultDir$, "SpriteCreator v2\SpriteCreator.bas")
#lablog "if SpriteCreator .bas exists heading to [makeexe]............"
     if res then spritecreated = 1 : goto [makeEXE]

[preferences]
  confirm "No Preferences page yet"+chr$(13)+chr$(13)+"Was this helpful?";lol$
 wait

' a program to select a bas file to get it's Line count
[numofLines]
#lablog "@- [numofLines]............"
  dim line$(20000)
  filedialog "Open \ Select a Liberty Basic Source File (.bas) ", DefaultDir$; "\*.bas", file2Check$
     if file2Check$ = "" then wait
    open file2Check$ for input as #1
  while eof(#1) = 0
      line input #1, line$(x)
      spaces$ = line$(x)
      if spaces$ = "" then y = y + 1
      x = x +1
  wend
 close #1
 file2Check$ = GetFilename$(file2Check$)
  message$ = chr$(13);chr$(13);"   ";file2Check$;chr$(13);chr$(13);"   ";x-y;" lines of code";chr$(13);"   ";y;"  lines of spaces ";chr$(13);"   ";x;"  lines in total";chr$(13);"    Approx ";x/25;" pages"
  message$ = GetMessage$(message$)
wait

'subroutine for selections of combo boxes
sub asciiSelected asciiList$
#lablog "entering sub asciiSelected asciiList$..........."
    #main.asciiList, "selection? asciiChoice$"
end sub

sub lbsampleSelected lbsamplesList$
#lablog "entering sub lbsampleSelected lbsamplesList$..........."
  q$ = chr$(34)
  call saveValue
    print "lbSamplesPath = ";upath$;"\AppData\Roaming\Liberty Basic v4.5.1\";lbsamps$
    #main.lbsamplesList, "selection? lbsamps$"
    lbsamps$ = lbsamps$;".bas"
    runFile$ = upath$;"\AppData\Roaming\Liberty Basic v4.5.1\";lbsamps$
    print "lbSamplesPath = ";upath$;"\AppData\Roaming\Liberty Basic v4.5.1\";lbsamps$
    lbRunIt$ = lbpath$;"\";lbexe$
  if fileExists(upath$;"\AppData\Roaming\Liberty Basic v4.5.1", lbsamps$) <> 0 then
        runFile$ = upath$;"\Application Data\Liberty Basic v4.5.1\";lbsamps$
         run q$;lbRunIt$;q$;" ";q$;runFile$;q$
  else
        if fileExists(upath$;"\Application Data\Liberty Basic v4.5.1", lbsamps$) <> 0 then
          runFile$ = upath$;"\Application Data\Liberty Basic v4.5.1\";lbsamps$
          run q$;lbRunIt$;q$;" ";q$;runFile$;q$
       else
          notice "Can't find ";lbsamps$;".txt"
      end if
  end if
 end sub

 sub lbdialogSelected lbdialogsList$
#lablog," entering sub lbdialogSelected lbdialogsList$......."
 q$ = chr$(34)
   call saveValue
#main.lbdialogsList, "selection? lbdialog$"
    if lbdialog$ = " Color Dialog" then runFile$ = lbpath$;"\lb4help\LibertyBASIC_4_web\html\libe00mf.htm"
    if lbdialog$ = " Printer Dialog" then runFile$ = lbpath$;"\lb4help\LibertyBASIC_4_web\html\libe72sp.htm"
    if lbdialog$ = " File Dialog" then runFile$ = lbpath$;"\lb4help\LibertyBASIC_4_web\html\libe7ezo.htm"
    if lbdialog$ = " All lbDialogs" then runFile$ = lbpath$;"\lb4help\LibertyBASIC_4_web\html\libe6gmr.htm"
    if lbdialog$ = " Font Dialog" then runFile$ = lbpath$;"\lb4help\LibertyBASIC_4_web\html\libe7sc5.htm"
    if lbdialog$ = " Confirm Dialog" then runFile$ = lbpath$;"\lb4help\LibertyBASIC_4_web\html\libe4pps.htm"
    if lbdialog$ = " Notice Dialog" then runFile$ = lbpath$;"\lb4help\LibertyBASIC_4_web\html\libe1836.htm"
   if lbdialog$ = " Prompt Dialog" then runFile$ = lbpath$;"\lb4help\LibertyBASIC_4_web\html\libe2p4c.htm"
   print "lbDialogselected =";runFile$ 'for testing with mainwin
if fileExists(DefaultDir$,"htmlviewer.exe") <> 0 then
         run "htmlviewer.exe ";runFile$
     else
         run "explorer.exe ";runFile$
   end if
 end sub

 sub lbreservedwordSelected lbreservedwordList$
#lablog," entering sub lbreservedwordSelected lbreservedwordList$ ........."
   call saveValue
   #main.lbdialogsList, "selection? lbreserved$"
 end sub

sub lbbakfileSelected lbbakfilesList$
#lablog," entering sub lbbakfileSelected lbbakfilesList$ "
q$ = chr$(34)
  call saveValue
    #main.lbbakfilesList, "selection? lbbakfile$"
      lbbakfile$ = lbbakfile$;".bak"
      print "lbbakfilePath = ";upath$;"\AppData\Roaming\Liberty BASIC v4.5.1\bak";lbbakfile$ 
  if fileExists(upath$;"\AppData\Roaming\Liberty BASIC v4.5.1\bak", lbbakfile$) <> 0 then
        runFile$ = upath$;"\AppData\Roaming\Liberty BASIC v4.5.1\bak\";lbbakfile$
        lbRunIt$ = lbpath$;"\";lbexe$
        run q$;lbRunIt$;q$;" ";q$;runFile$;q$
  else
       if fileExists(upath$;"\Application Data\Liberty BASIC v4.5.1\bak\", lbbakfile$) <> 0 then
         runFile$ = upath$;"\Application Data\Liberty BASIC v4.5.1\bak\";lbbakfile$
         lbRunIt$ = lbpath$;"\";lbexe$
         run q$;lbRunIt$;q$;" ";q$;runFile$;q$
      else
         notice "can't find ";lbbakfile$
      end if
  end if
 end sub

'next few subroutines to GET the info to populate the combo boxes
 sub getAscii
 #lablog," entering sub getAscii"
 dim asciiList$(250)
 y = 7
   asciiList$(0)= "     Controls"
   asciiList$(1) = " chr$(0) = (nul) ";chr$(0)
   asciiList$(2) = " chr$(27) = (escape) ";chr$(27)
   asciiList$(3) = " chr$(32) = (space) ";chr$(32)
   asciiList$(4) = " chr$(13) = (enter) ";chr$(13)
   asciiList$(5) = "      Printables"
   asciiList$(6) = " chr$(32)= (space) ";chr$(32)
 for x = 33 to 255
 print "asciiList$(y) =  ";"chr$(";x;") = ";chr$(x)
   asciiList$(y) = " chr$(";x;") = ";chr$(x)
    y = y + 1
 next x
   #main.asciiList, "reload"
 end sub

sub getlbsamples
#lablog," entering sub getlbsamples "
  q$ = chr$(34)
   dim folderInfo$(1, 1)
   dim lbsamplesList$(10)
    files upath$;"\Application Data\Liberty BASIC v4.5.1\", folderInfo$()
    numberoFiles = val(folderInfo$(0, 0))
   redim lbsamplesList$(numberoFiles)
 for x = 0 to numberoFiles
    print folderInfo$(x, 0)
     filename$ = folderInfo$(x, 0)
    if right$(filename$, 3) <> "bas" and right$(filename$, 3) <> "BAS" then [discardThisLine]
        lbsamplesList$(x) = left$(filename$, len(filename$) - 4)
     print "folderInfo$(x, 0) = ";folderInfo$(x, 0)
     print "  lbsamplesList$(x) = ";lbsamplesList$(x)
    [discardThisLine]
 next x
   sort lbsamplesList$(), 0 ,numberoFiles
  #main.lbsamplesList, "reload"
end sub

sub getlbdialogs
#lablog," entering sub getlbdialogs"
lbFontDialog$ = " All lbDialogs, Prompt Dialog, Notice Dialog, Font Dialog, Color Dialog, File Dialog, Printer Dialog, Confirm Dialog"
 for x = 1 to 8
    print "filename$ = ";word$(lbFontDialog$, x, ",")
    filename$ = word$(lbFontDialog$, x, ",")
  lbdialogsList$(x) = filename$
    print "lbdialogsList$(x) = ";lbdialogsList$(x)
 next x
   sort lbdialogsList$(), 0 ,8
  #main.lbdialogsList, "reload"
end sub

 sub getlbreservedwords
 #lablog," entering sub getlbreservedwords "
       q$ = chr$(34)
    dim lbreservedwordsList$(300)
  for x = 0 to 300
       filename$ = word$(lbReservedWords$, x ,",")
        print "filename$ = ";filename$
       lbreservedwordsList$(x) = filename$
       print "lbreservedwordsList$ = ";lbreservedwordsList$(x)
  next x
     sort lbreservedwordsList$(), 1 ,300
     #main.lbreservedwordsList, "reload"
 end sub

sub getlbBakFiles
 #lablog," entering sub getlbBakFiles "
    dim folderInfo$(1, 1)
    dim lbbakfilesList$(500)
    files upath$;"\Application Data\Liberty BASIC v4.5.1\bak", folderInfo$()
    numberOfFiles = val(folderInfo$(0, 0))
     redim lbbakfilesList$(numberOfFiles)
  for x = 0 to numberOfFiles
       print folderInfo$(x, 0)
       filename$ = folderInfo$(x, 0)
      if right$(filename$, 3) <> "bak" then [skip]
       lbbakfilesList$(x) = left$(filename$, len(filename$) - 4)
       print "folderInfo$(x, 0) = ";folderInfo$(x, 0)
       print "  lbbakfilesList(x) = ";lbbakfilesList$(x)
 [skip]
  Next x
     sort lbbakfilesList$(), 0 ,numberOfFiles
    #main.lbbakfilesList, "reload"
    #lablog, "got lbBakFiles List......"
end sub

 sub resetRadioOptions
  #lablog," entering sub resetRadioOptions"
    dictionary$ = "" : keyCount = 0 : lastKey$ = "" : selectedKey$ = ""
    call readDictionary
    call loadKeys
    #main.value, "!origin 0, 0 "
    #main.value, "!setfocus"
 end sub

'subroutine to GET the current Users HomePath
 sub getUserPath
   #lablog," entering sub getUserPath"
     cursor hourglass
     run "cmd.exe /c echo %userprofile% |clip", HIDE
  call pause 3000
     open "GetUserPath" for text as #1
     #1 "!paste"
     #1 "!contents? upath$"
     upath$ = trim$(upath$)
     print "upath$ =  ";upath$ 'print for testing with mainwin
     close  #1
   cursor normal
   #lablog, "Got user path - Exiting  sub getUserPath......"
 end sub

 'create a project and tkn file and add it to the MyProjects List
[makeproject]
#lablog," @ - [makeproject]"
     if categorie$ <> MyProjects$ then
        notice "You must first select Radio Button >> MyProjects" : wait
     end if
       tkn = 2
       project = 1
  goto [bas2exe]

[remakeproject]
 #lablog," @ - [remakeproject]"
     if selectedKey$ = "" then notice "No Listing selected. Select an item from "+categorie$+ " list and Try Again " : wait
     if fileExists(savedProjects$;"\";selectedKey$,selectedKey$;".bas") = 0 then notice selectedKey$+chr$(13)+" of "+categorie$+" wasn't saved"+chr$(13)+"Try [Make New Project], BAS<2>EXE, or BAS<2>TKN"+chr$(13)+"And Select the appropriate .bas file" : wait
    project = 1
    tkn = 4
     fname$ = savedProjects$;"\";selectedKey$;"\";selectedKey$;".bas"
  goto [bas2exe]

[bas2tkn]
     if categorie$ <> programs$ then notice "You must first select Radio Button >> Programs" : wait
      #lablog," @ - [bas2tkn]"
       tkn = 3
goto [bas2exe]

[makeEXE]
          if categorie$ <> programs$ then
             notice "You must first select Radio Button >> Programs" : wait
          end if
          tkn = 0
 'BAS2EXE Version v1.8a For Linux/WINE,  Windows 10 (possibly XP, Win 7, 8)
'Date = July 2021
'Title - BAS2EXE v1.8
'Author - xxgeek, a member of the justbasiccom.proboards.com/ forums
 print "Starting into BAS2EXE"
[bas2exe]
#lablog, "@ - [bas2exe] - calling fixtime, and fixdate"
call fixtime
call fixdate

if tkn = 0 then print "@ [bas2exe] Starting - Running BAS<2>EXE user chooses if full project, plus adding new listing in Programs"
if tkn = 2 then print "@ [bas2exe] Starting  - Making New Project, plus creating new listing in MyProjects category"
if tkn = 3 then print "@ [bas2exe] Starting  - Making TKN, plus adding listing in Programs category"
if tkn = 4 then print "@ [bas2exe] Starting to remake project ";selectedKey$;" ReWriting MyProjects Listing"

[go]
#lablog," @ - [go]  - Starting to make window for make tkn, bas2exe, or makeproject"
 print "@ [go] - Starting to make window"
 print " "
  if tkn = 4 then [spriteOnly]
  if spritecreated = 1 then [spriteOnly]

 ' setup a Window for User to Select a .bas File to Make a Project with
  'nomainwin
    WindowWidth = 600
    WindowHeight = 380
    UpperLeftX=INT((DisplayWidth-WindowWidth)/2)
    UpperLeftY=INT((DisplayHeight-WindowHeight)/2)
    BackgroundColor$ = "lightgray"
    ForegroundColor$ = "black"

'add some text ,some buttons, and checkboxes to the Window
    statictext #pick.header, "   BAS <2> EXE", 165, 20, 590, 45
    statictext #pick.exe, "EXE FILE", 15, 70, 105, 30
    statictext #pick.temp, "Temp Files", 390, 70, 105, 30
    statictext #pick.warning, "Test in Liberty Basic before Proceeding", 190, 310, 260, 15
    statictext #pick.info, "Select a working Liberty Basic Source Code File (.bas)", 30, 220, 590, 30
    statictext #pick.lbforums, "Visit the Liberty Basic Forums @ https://libertybasiccom.proboards.com/", 90, 335, 590, 20

    checkbox #pick.password, "Passworded", [yespass], [nopass], 20, 155, 140, 20
    checkbox #pick.bit32, "32 Bit", [bit32], [nobit32], 20, 105, 80, 20
    checkbox #pick.bit64, "64 Bit", [bit64], [nobit64], 20, 130, 80, 20
    checkbox #pick.incbas, "Don't Include Source", [noincsource], [incsource], 20, 180, 180, 20
    checkbox #pick.sed, "Keep SED", [sed], [nosed], 400, 105, 140, 20
    checkbox #pick.vbs, "Keep VBS", [keepvbs], [novbs], 400, 130, 140, 20
    checkbox #pick.project, "Don't Keep Project Dir", [noproject], [project], 400, 155, 160, 20

    button #pick.default, "Select File", [defaultClick],UL 140, 270, 135, 35
    button #pick.32, "Cancel", [cancel],UL 320, 270, 135, 35

'open the Window, and set some Fonts for each statictext, and buttons
#lablog, "Opening Select File window........"
open "BAS2EXE v1.8" for window_nf as #pick
 #pick, "trapclose [quit.pick]"
 #pick, "font Arial 10 bold"
 #pick.header, "!font Arial 24 bold"
 #pick.exe, "!font Arial 14 bold"
 #pick.temp, "!font Arial 14 bold"
 #pick.sed, "font Arial 12 bold"
 #pick.vbs, "font Arial 12 bold"
 #pick.project, "font Arial 12 bold"
 #pick.password, "font Arial 12 bold"
 #pick.info, "!font Arial 18 bold"
 #pick.lbforums "!font Arial 10 bold"
 #pick.32, "!font Arial 12 bold"
 #pick.bit64, "font Arial 12 bold"
 #pick.bit32, "font Arial 12 bold"
 #pick.incbas, "font Arial 12 bold"
 #pick.default, "!font Arial 12 bold"
 #pick.warning, "!font Arial 8 bold"
 #pick.default, "!setfocus"
  print "window up and running "
#lablog, "Select File window opened......."
  #lablog$, "If tkn = 3 then BAS<2>EXE Button was pressed. Creating Make New Project Window"
   if tkn = 3 then
       #pick.temp, "!HIDE"
       #pick.exe "!HIDE"
       #pick.header "BAS < 2 > TKN"
       #pick.bit64, "HIDE"
       #pick.bit32, "HIDE"
       #pick.sed, "HIDE"
       #pick.vbs, "HIDE"
   end if

   if tkn = 2 then
         #lablog$, "If tkn = 2 then Make Project Button was pressed. Creating Make New Project Window"
         #pick.project, "HIDE"
         #pick.temp, "!HIDE"
         #pick.incbas, "HIDE"
         #pick.exe "!HIDE"
         #pick.header "Make New project"
         #pick.bit64, "HIDE"
         #pick.bit32, "HIDE"
         #pick.sed, "HIDE"
         #pick.vbs, "HIDE"
        #pick.password, "HIDE"
   end if
wait

 [incsource]
#lablog," @ - [incsource]  - include source checkbox selected"
   incbas = 1
 wait
 [noincsource]
 #lablog," @ - [noincsource]  - include source checkbox deselected"
  incbas = 0
 wait
 [sed]
 #lablog," @ - [sed] "
  sed = 1
 wait
 [nosed]
 #lablog," @ - [nosed] "
  sed = 0
 wait
 [keepvbs]
  #lablog," @ - [keepvbs] "
  vbs = 1
 wait
 [novbs]
  #lablog," @ - [novbs] "
  vbs = 0
 wait
 [project]
  #lablog," @ - [project] "
  project = 1
  wait
 [noproject]
  #lablog," @ - [noproject] "
  project = 0
 wait
 [yesTKN]
  #lablog," @ - [yesTKN] "
 tkn = 0
 wait
  [noTKN]
    #lablog," @ - [noTKN] "
 tkn = 1
 wait
  [yesBAS]
    #lablog," @ - [yesBAS] "
 bas=0
 wait
  [noBAS]
    #lablog," @ - [noBas] "
 bas = 1
 wait

' passworded exe is true(user selected)
[yespass]
  #lablog," @ - [yespass] "
 p=1
 wait
'passworded exe is false, default
[nopass]
  #lablog," @ - [nopass] "
 p=0
 wait

'make 32 bit exe = true(user selected)
[bit32]
  #lablog," @ - [bit32] "
bit=32
#pick.bit64, "hide"
wait

'make 64 bit exe, default
[bit64]
  #lablog," @ - [bit64] "
bit=64
#pick.bit32, "hide"
wait

[nobit32]
  #lablog," @ - [nobit32] "
bit=64
#pick.bit64, "show"
wait

[nobit64]
  #lablog," @ - [nobit64] "
  bit=0
  #pick.bit32, "show"
wait

'close the open window for Selecting bas file
[defaultClick]
cursor hourglass
#lablog, "Select File button pressed - closing Select Source File window "
 print "Select File button pressed - closing Select Source File window "
  close #pick

'check existence and lbPath$ (Liberty Basic default install dir)
#lablog, "checking path existence for ";lbpath$
res = pathExists(lbpath$)
  if res then a = a + 1 else notice " Liberty Basic v4.5.1 was not was not found in ";lbpath$;"  - aborting mission":  wait

[spriteOnly]
print "@ [spriteOnly]"
  #lablog," @ - [spriteOnly] "
' Liberty Basic 4.5.1 is installed - continue on
 'define some variables
  'p=0 'passworded exe = false
  'lbexe$ = "liberty.exe"
  'lbruntime$ = "lbrun2.exe"
  DllList$="vbas31w.sll vgui31w.sll voflr31w.sll vthk31w.dll vtk1631w.dll vtk3231w.dll vvm31w.dll vvmt31w.dll"
  savedProjects$ = "savedProjects"
#lablog, "Checking existence for lbrun$ lbexe$ lbpath$, and all the supproting dll and sll files."
 'Checking all paths and file locations for existence (dll's, sll's, lbasic.exe, and lbrun2.exe)
#lablog, "Checking lbexe$"
res=fileExists(lbpath$, lbexe$)
   if res then a = a + 1 else notice lbexe$;" lbexe$ Does not exist in  ";lbpath$;"  - aborting mission":  wait
#lablog, "Checking lbruntime$"
res=fileExists(lbpath$,lbruntime$)
    if res then a = a + 1 else notice lbrun$;" lbrun$ Does not exist in  ";lbpath$;"  - aborting mission":  wait
#lablog, "Checking vbas31w.sll"
res=fileExists(lbpath$,"vbas31w.sll")
    if res then a = a + 1 else notice " vbas31w.sll Does not exist in  ";lbpath$;"  - aborting mission":  wait
#lablog, "Checking vgui31w.sll"
res=fileExists(lbpath$,"vgui31w.sll")
    if res then a = a + 1 else notice " vgui31w.sll Does not exist in  ";lbpath$;"  - aborting mission":  wait
#lablog, "Checking voflr31w.sll"
res=fileExists(lbpath$,"voflr31w.sll")
    if res then a = a + 1 else notice " voflr31w.sll Does not exist in  ";lbpath$;"  - aborting mission":  wait
#lablog, "Checking vthk31w.dll"
res=fileExists(lbpath$,"vthk31w.dll")
    if res then a = a + 1 else notice " vthk31w.dll Does not exist in  ";lbpath$;"  - aborting mission":  wait
#lablog, "Checking vtk1631w.dll"
res=fileExists(lbpath$,"vtk1631w.dll")
    if res then a = a + 1 else notice " vtk1631w.dll Does not exist in  ";lbpath$;"  - aborting mission":  wait
#lablog, "Checking vtk3231w.dll"
res=fileExists(lbpath$,"vtk3231w.dll")
    if res then a = a + 1 else notice " vtk3231w.dll Does not exist in  ";lbpath$;"  - aborting mission":  wait
#lablog, "Checking vvm31w.dll"
res=fileExists(lbpath$,"vvm31w.dll")
    if res then a = a + 1 else notice " vvm31w.dll Does not exist in  ";lbpath$;"  - aborting mission":  wait
#lablog, "Checking vvmt31w.dll"
res=fileExists(lbpath$,"vvmt31w.dll")
    if res then a = a + 1 else notice " vvmt31w.dll Does not exist in  ";lbpath$;"  - aborting mission":  wait
  #lablog," all support files dll's sll's lbrun2.exe, and lbasic.exe accounted for "
' all runtime support files accounted for

'prompt user for a password to start the created EXE File
   if p=0 then [filediag]
   #lablog, "Prompting user for password to add to the exe file about to be created"
   if p= 1 then Prompt "TYPE a PASSWORD"+chr$(13)+ "Password for EXE file is:       (no spaces)";passwerd$
  if passwerd$ = "" then notice "BAS2EXE will continue, without placing a password on the EXE file created" : p = 0

' Use the filedialog function to allow user to select a source file (.bas)
[filediag]
print "Opening FileDialog - User chooses .bas file to create TKN, EXE, or Project"
  if spritecreated = 1 then fname$ = DefaultDir$;"\SpriteCreator v2\SpriteCreator.bas" : goto [spriteOnly2]
  if tkn = 4 then [spriteOnly2]
print "Opening FileDialog - User chooses .bas file to create TKN, EXE, or Project"
#lablog, "Opening FileDialog - User chooses .bas file to create TKN, EXE, or Project"
'open file dialog to choose a .bas file for exe conversion
 filedialog "Open \ Select a Liberty Basic Source File (.bas) ", DefaultDir$; "\*.bas", fname$
     if fname$ = "" then notice "No file selected" :  wait

[spriteOnly2] 'to make sure the lb support files are in the SpriteCreator v2 folder
print "file chosen = ";fname$
#lablog, "file chosen = ";fname$
print "separating name from path, and name from extension"
#lablog, "separating name from path, and name from extension"

 'Separate path from selected filename, and extension from selected filename
  for var1 = len(fname$) to 1 step -1
     if mid$(fname$, var1, 1) = "\" then var2 = var1 -1 : var3 = var2 - ((len(fname$))) : exit for
  next var1
   var3 = abs(var3)
   orig$ = left$(fname$, var2)
   fname0$ = right$(fname$, var3 -1)
 for var4 = len(fname0$) to 1 step -1
    if mid$(fname0$, var4, 1) = "." then var5 = var4 -1 : var6 = var5 - ((len(fname0$))) : exit for
 next var4
 var6 = abs(var6)
 fnamenobas$ = left$(fname0$, var5)
  #lablog, "fnamenobas$ = ";fnamenobas$
 ' fname$ = Full Path of User Selected .bas file (including the filename.bas)
 ' fname0$ = Name of the Selected .bas File Only - eg ; filename.bas
 ' fnamenobas$ = Name of the Selected .bas File (without the .bas) - eg: filename
print "finished separating........."
[begin]
print "@ [begin] creating folders for projects, exe'e, sed's, vbs, and tkn files"
#lablog, "@ [begin] creating folders for projects, exe'e, sed's, vbs, and tkn files"
'define Destpath1$ as lb Projects\Current Project Folder
 DestPath$=DefaultDir$ 'Where this file is RUN from
 DestPathU$ = DestPath$;"\";savedProjects$ 'Projects Folder
 DestPath1$=DestPathU$;"\";fnamenobas$ 'Current created Project Folder

'Make Folders for Liberty Basic Projects, EXE files, TKN files, BAS files, SED files and Current Projects
  res = mkdir(DestPathU$) 'projects dir
  res = mkdir(DestPath1$) 'new project dir = name of selected bas file (no .bas) in Projects Dir
  res =mkdir(DefaultDir$;"\";"EXE") 'exe files saved here
  res = mkdir(DefaultDir$;"\";"TKN") 'tkn files saved here
  res= mkdir(DefaultDir$;"\";"BAS")  'selected bas file saved here (includes password code if exe was passworded)
  res= mkdir(DefaultDir$;"\";"SED") ' saves the created SED file (self extracting directive)
  res= mkdir(DefaultDir$;"\";"VBS") ' saves VBS file (.vbs script that auto clicks `save tkn`, and `saved as` buttons)

'make sure Folders were actually created
#lablog, "Verifying new folders in ";DestPath$;" exist......"
 res=pathExists(DestPathU$)
   if res then a=a+1 else notice "savedProjects folder was NOT Created in ";DestPath$;"  - aborting mission":  wait
 res=pathExists(DestPath1$)
      if res then a=a+1 else notice "New Folder ";fnamenobas$;" was NOT Created in ";DestPath1$;"  - aborting mission":  wait
tknFolder$=DefaultDir$;"\";"TKN"
res=pathExists(tknFolder$)
   if res then a=a+1 else notice "TKN Folder was NOT Created in ";DestPath$;"  - aborting mission":  wait
   basFolder$=DefaultDir$;"\";"BAS"
res=pathExists(basFolder$)
   if res then a=a+1 else notice "BAS Folder was NOT Created in ";DestPath$;"  - aborting mission":  wait
print "folders created, and verified........."
#lablog, "folders created, and verified........."

if spritecreated = 1 or tkn = 4 then [noDelete]

'remove existing fname0$ from dir before creating new one
print "removing existing same name file prior to copying bas file to new project dir"
#lablog, "removing existing same name file prior to copying bas file to new project dir"
   if fileExists(DestPath1$, fname0$) <> 0 then kill DestPath1$;"\";fname0$

[noDelete]
'copy selected bas file to Projects\current project folder
 q$= chr$(34)
 if spritecreated = 1 then DestPath1$ = DefaultDir$;"\SpriteCreator v2"
print "copying selected bas file to Projects\current project folder"
#lablog, "copying selected bas file to Projects\current project folder"
  open fname$ for input as #fname
 fnameTemp$="tempBas.bas"
  open fnameTemp$ for output as #2

'add a password prompt to the begining of the temp bas file(to be added to the exe)
 if p=0 then [nopasswerd] '
 #lablog, "Prompting user for password to temp bas file used for EXE file creation............"
     print "Prompting user for password to temp bas file used for EXE file creation"
     #2, "prompt ";q$;"Enter the Password";q$;";";"passwerd$"
     #2, "if passwerd$ <> ";q$;passwerd$;q$;" then end"
     print "password appended to temp bas file            "
     #lablog, "password appended to temp bas file for password to temp bas file used for EXE file creation............"
  [nopasswerd]
    #2, input$(#fname, lof(#fname));
    close #fname
    close #2
#lablog, "copying ";fnameTemp$;" to current project folder";DestPath$;" for input "
print "opening fnameTemp$ for input"

'copy temp.bas file to current project folder
 open fnameTemp$ for input as #1
 if fileExists(DestPath1$, fname0$ ) <> 0 then kill DestPath1$;"\";fname0$
 open DestPath1$;"\";fname0$ for output as #2
   print #2, input$(#1, lof(#1));
  close #2
  close #1

'check if the current project .bas file was copied to new dir
res=fileExists(DestPath1$,fname0$)
    if res then a = a + 1 else notice fname0$; " Was not copied to  ";DestPath1$;"  - aborting mission":  wait
print "finished copying and verifying bas file exists new project dir........"
#lablog, "finished copying and verifying bas file exists new project dir........"

'copy selected .bas file to BAS dir and date it
print "start copying bas file to BAS dir and dating it........"
#lablog, "start copying bas file to BAS dir and dating it........"
 open DestPath1$;"\";fname0$ for input as #file
  open DestPath$;"\";"BAS\";fnamenobas$;fixeddate$;".bas" for output as #1
     print #1, input$(#file, lof(#file));
 close #file
  close #1

'remove any existing exe of same name as bas file selected only if created on same date
print "remove any existing exe of same name as bas file selected only if created on same date."
#lablog, "removing any existing exe of same name as bas file selected only if created on same date"
if fileExists(DestPath$;"\EXE", fnamenobas$;fixeddate$;".exe.BAK")  <> 0 then kill DestPath$;"\";"EXE";"\";fnamenobas$;fixeddate$;".exe.BAK"
    print "finished checking and deleting existing bas file in BAS dir..........."
    print "starting copying necessary dll, and sll files to new project dir......."

  if tkn = 4 then [verifyDLLs]
#lablog, "If tkn = 4, jumping support file creation(should already exist"
 'Copy the needed DLL and SLL files from Liberty Basic dir to projects\projectname Dir
 print "Copying the needed DLL and SLL files from Liberty Basic dir to projects\projectname Dir"
#lablog, "Copying DLL and SLL files from Liberty Basic dir to projects\projectname Dir"
 runtimeSupportFile$ = ""
 i = 0
 while 1
    i = i + 1
    runtimeSupportFile$=word$(DllList$,i)
  if runtimeSupportFile$ ="" then exit while
    sourceFile$=lbpath$;"\";runtimeSupportFile$
    destinationFile$=DestPath1$;"\";runtimeSupportFile$
  if spritecreated = 1 then destinationFile$ = DefaultDir$;"\SpriteCreator v2\";runtimeSupportFile$

'remove any existing lb runtime support files from files to copy
#lablog, "Copying ";sourceFile$;" to ";destinationFile$
   print  "Copying ";sourceFile$;" to ";destinationFile$
if fileExists(DestPath1$, runtimeSupportFile$) <> 0 then  [fileExists2]
  open sourceFile$ for input as #file
  open destinationFile$ for output as #1
  print #1, input$(#file, lof(#file));
  close #file
  close #1
  [fileExists2]
 wend

[verifyDLLs]
'verify dll's and sll's were copied to temp folder
#lablog, "@ -[verify DLLs] - verifying dll and sll files were copied to ....  ";DestPath1$
  res=fileExists(DestPath1$,"vbas31w.sll")
    if res then a = a + 1 else notice " vbas31w.sll Was not created in -->  ";DestPath1$;"  - aborting mission":  wait
    if res then a = a + 1 else notice " vgui31w.sll Was not created in -->  ";DestPath1$;"  - aborting mission":  wait
res=fileExists(DestPath1$,"voflr31w.sll")
    if res then a = a + 1 else notice " voflr31w.sll Was not created in -->  ";DestPath1$;" - aborting mission":  wait
res=fileExists(DestPath1$,"vthk31w.dll")
    if res then a = a + 1 else notice " vthk31w.dll Was not created in -->  ";DestPath1$;" - aborting mission":  wait
res=fileExists(DestPath1$,"vtk1631w.dll")
    if res then a = a + 1 else notice " vtk1631w.dll Was not created in -->  ";DestPath1$;"  - aborting mission":  wait
res=fileExists(DestPath1$,"vtk3231w.dll")
    if res then a = a + 1 else notice " vtk3231w.dll Was not created in -->  ";DestPath1$;"  - aborting mission":  wait
res=fileExists(DestPath1$,"vvm31w.dll")
    if res then a = a + 1 else notice " vvm31w.dll Was not created in -->  ";DestPath1$;"  - aborting mission":  wait
res=fileExists(DestPath1$,"vvmt31w.dll")
                if res then a = a + 1 else notice " vvmt31w.dll Was not created in -->  ";DestPath1$;"  - aborting mission":  wait
#lablog, "support files .dlls and .slls verified in ";DestPath1$
   print  "support files verified............"

'remove any left over existing lbrun2.exe (created by errors) from new project before creating new one
'Liberty Basic can't create\rename a file that exists, so if it does already exist - kill it (delete it)
#lablog, "removing any left over existing run451.exe "
   print  "removing any left over existing run451.exe "
   if fileExists(DestPath1$, lbruntime$) <> 0 then kill DestPath1$;"\"; lbruntime$

'copy lbrun2.exe to Current Project Folder
#lablog, "coping lbrun2.exe to Current Project Folder"
   print  "copying lbrun2.exe to Current Project Folder "
 open lbpath$;"\";lbruntime$ for input as #file
 open DestPath1$;"\";lbruntime$ for output as #1
  print #1, input$(#file, lof(#file));
  close #file
  close #1

'rename lbrun2.exe to name of User Selected .bas File - .bas +.exe
#lablog, "renaming lbrun2.exe to name of User Selected .bas File - .bas +.exe"
   print  "renaming lbrun2.exe to name of User Selected .bas File - .bas +.exe "
   if fileExists(DestPath1$, fnamenobas$;".exe") <> 0 then kill DestPath1$;"\";fnamenobas$;".exe"
    name DestPath1$;"\";lbruntime$ as DestPath1$;"\";fnamenobas$;".exe"

'check new exe (renamed lbrun2.exe) file for existence in current project Folder )
#lablog, "checking new exe (renamed lbrun2.exe) file for existence in current project Folder )"
   print  "checking new exe (renamed lbrun2.exe) file for existence in current project Folder ) "
res=fileExists(DestPath1$,fnamenobas$;".exe")
  if res then a=a+1 else notice "lbrun2.exe not copied or renamed - aborting mission":  wait
print "finished deleting, and verifying the rename of lbrun2 in new project dir"
#lablog, "lbrun2.exe was renamed to ";selectedKey$;".exe"

'check for any left over tkn file existence delete if true
#lablog, "checking for any left over tkn file existence due to errors on past runs, if true, deleting ";fnamenobas$;".tkn"
   print  "checking for any left over tkn file, deleting ";fnamenobas$;".tkn";" if true"
   if fileExists(DestPath1$, fnamenobas$;".tkn") <> 0 then kill DestPath1$;"\";fnamenobas$;".tkn"
#lablog, "left over tkn file existence verified and if true, deleted"
   print  "left over tkn file existence verified and if true, deleted"

'#######################################################################
'Write Visual Basic Script to a .vbs file to automate TKN "save" and "information"
autoSave$ = "autoSave.vbs"
  open autoSave$ for output as #1
  #lablog, "writing to disk autosave.vbs script to automate tkn creation/save events"
  print #1, "Set WshShell = WScript.CreateObject(";q$;"WScript.Shell";q$;")"
 #1, "WshShell.AppActivate ";q$;"Save *.TKN File As...";q$
 '#1, "Wscript.Sleep(200)" - keeping for testing
 #1, "WshShell.SendKeys ";q$;"{ENTER}";q$
 #1, "Wscript.Sleep(550)" 'this delay may need adjusting on your pc
 #1, "WshShell.AppActivate ";q$;"saved as";q$
 #1, "WshShell.SendKeys ";q$;"{ENTER}";q$
  close #1
#lablog, "script written........"
   print  "script written........"

'loop until autoSave$ File is verified
#lablog, "looping until autoSave$ File is verified.............."
   print  "looping until autoSave$ File is verified..............."
 do
 res = fileExists(DestPath$,autoSave$)
  if res then exit do
   scan
 loop until res
print "autosave vbs file written, saved, and existence verified............."
#lablog, "autosave vbs file written, saved, and existence verified............."

'#######################################################################

'Create the TKN file in Projects\current project folder.
print "creating the tkn file..........."
#lablog, "creating the tkn file..........."
  RUN lbpath$;"\";lbexe$;" -T -A ";DestPath1$;"\";fname0$
'give time for the save TKN window to appear
  call pause 1500

'#######################################################################
 'run the script to close the "save" dialog, and the follow up notice of creation automatically
   #lablog," running autoSave vbs script to auto 'click' ENTER on 'save as' dialog and Information dialog "
  run "wscript  ";autoSave$
'#######################################################################
call pause 1500

'loop until TKN File is verified saved
#lablog, "verifying tkn file created..........."
 do
 res = fileExists(DestPath1$,fnamenobas$;".tkn")
  if res then exit do
   scan
 loop until res
print "tkn verified saved to new project dir"
#lablog, "tkn verified and saved to ";DestPath1$;"\";fnamenobas$;".tkn"

'remove any existing tkn of same name in TKN dir
#lablog, "checking for existing";DestPath1$;"\TKN\"; fnamenobas$;fixeddate$;".tkn"
 if fileExists (DefaultDir$;"\TKN", fnamenobas$;fixeddate$;".tkn") <> 0 then kill DefaultDir$;"\TKN\";fnamenobas$;fixeddate$;".tkn"

'let liberty basic cool off for a second just to be nice :D
call pause 500

'copy TKN$ file to TKN dir, and date it
print "copying tkn to TKN dir and dating it............."
#lablog, "copying tkn to TKN dir and dating it............."
 open DestPath1$;"\";fnamenobas$;".tkn" for input as #file
 open DefaultDir$;"\TKN\";fnamenobas$;fixeddate$;".tkn" for output as #1
    print #1, input$(#file, lof(#file));
  close #file
  close #1

 if fileExists (DefaultDir$;"\TKN", fnamenobas$;fixeddate$;".tkn") <> 0 then
 print "tkn file veried saved to ";DefaultDir$;"\TKN"
#lablog, "tkn file veried saved to ";DefaultDir$;"\TKN"
      goto [continueOn]
  else
   notice fnamenobas$;fixeddate$;".tkn";" was NOT created in ";DefaultDir$;"\TKN" : goto [noiex]
 end if

[continueOn]
'check what tkn value =, and decide next step
print "@ [continueOn] - tkn file verified dated and saved to TKN dir..........."
#lablog, "@ [continueOn] - tkn file verified dated and saved to TKN dir..........."
   if tkn = 2 then
    print "sending to [newKey]/[continue] to add to ";categorie$;" List........"
 #lablog, "sending to [newKey]/[continue] to add to ";categorie$;" List........"
      newKey$ = fnamenobas$
      categorie$ = MyProjects$
     goto [continue]
  end if

   if tkn = 3 then
       print "sending to [newKey]/[continue] to add to ";categorie$;" List........  ";programs$
 #lablog, "sending to [newKey]/[continue] to add to ";categorie$;" List........  ";programs$
      newKey$ = fnamenobas$
      categorie$ = programs$
     goto [continue]
  end if

  if tkn = 4 then
    print "sending to [newKey]/[continue] to add to ";categorie$;" List........"
 #lablog, "sending to [newKey]/[continue] to add to ";categorie$;" List........"
      newKey$ = fnamenobas$
      categorie$ = MyProjects$
     goto [continue]
  end if

'bypass making the EXE file if SpriteCreator was selected (applies to first run only)
if spritecreator = 1 then [noiex]

[makeexe]
    print "@ [makeexe] checking for existence of IEXPRESS to make exe"
 #lablog, "@ [makeexe] checking for existence of IEXPRESS to make exe"
res=fileExists("c:\windows\system32","iexpress.exe")
    if res then [makeSed] else notice "Cannot find file --> iexpress.exe in c:\windows\system32"+chr$(13)+"Known issue for users of WINE in Linux"+chr$(13)+"Check WINE Tricks for Adding IExpress(after each update)"
 wait

'make the sed file for iexpress to read and create the (Self Extracting Directorate) exe file
[makeSed]
    print "@ [makeSed] to create SED file for use with IEXPRESS commandline"
 #lablog, "@ [makeSed] to create SED file for use with IEXPRESS commandline"
'can't write text to files that include quotes, so use the ascii characters so they will print without syntax errors
 q$=chr$(34) ' double quotes to be printed around files and paths in sed file text"
  sedfile$=fnamenobas$;".sed"
 open sedfile$ for output as #sed
   #sed "[Version]"
   #sed "Class=IEXPRESS"
   #sed "SEDVersion=3"
   #sed "[Options]"
   #sed "PackagePurpose=InstallApp"
   #sed "ShowInstallProgramWindow=1"
   #sed "HideExtractAnimation=1"
   #sed "UseLongFileName=1"
   #sed "InsideCompressed=0"
   #sed "CAB_FixedSize=0"
   #sed "CAB_ResvCodeSigning=0"
   #sed "RebootMode=N"
   #sed "InstallPrompt=%InstallPrompt%"
   #sed "DisplayLicense=%DisplayLicense%"
   #sed "FinishMessage=%FinishMessage%"
   #sed "TargetName=%TargetName%"
   #sed "FriendlyName=%FriendlyName%"
   #sed "AppLaunched=%AppLaunched%"
   #sed "PostInstallCmd=%PostInstallCmd%"
   #sed "AdminQuietInstCmd=%AdminQuietInstCmd%"
   #sed "UserQuietInstCmd=%UserQuietInstCmd%"
   #sed "SourceFiles=SourceFiles"
   #sed "[Strings]"
   #sed "InstallPrompt="
   #sed "DisplayLicense="
   #sed "FinishMessage="
 exe$=fnamenobas$;".exe"
   #sed "TargetName=";q$;DefaultDir$;"\EXE\";exe$;q$
   #sed "FriendlyName=";q$;fnamenobas$;q$
   #sed "AppLaunched=";q$;exe$;q$
   #sed "PostInstallCmd=<None>"
   #sed "AdminQuietInstCmd="
   #sed "UserQuietInstCmd="
   #sed "FILE0=";q$;exe$;q$
 sedtkn$=fnamenobas$;".tkn"
   #sed "FILE1=";q$;sedtkn$;q$
  sll1$="vbas31w.sll"
  sll2$="vgui31w.sll"
  sll3$="voflr31w.sll"
  dll1$="vthk31w.dll"
  dll2$="vtk1631w.dll"
  dll3$="vtk3231w.dll"
  dll4$="vvm31w.dll"
  dll5$="vvmt31w.dll"
   #sed "FILE2=";q$;sll1$;q$
   #sed "FILE3=";q$;sll2$;q$
   #sed "FILE4=";q$;sll3$;q$
   #sed "FILE5=";q$;dll1$;q$
   #sed "FILE6=";q$;dll2$;q$
   #sed "FILE7=";q$;dll3$;q$
   #sed "FILE8=";q$;dll4$;q$
   #sed "FILE9=";q$;dll5$;q$
   #sed "[SourceFiles]"
   #sed "SourceFiles0=";q$;DestPath1$;q$
   #sed "[SourceFiles0]"
   #sed "%FILE0%="
   #sed "%FILE1%="
   #sed "%FILE2%="
   #sed "%FILE3%="
   #sed "%FILE4%="
   #sed "%FILE5%="
   #sed "%FILE6%="
   #sed "%FILE7%="
   #sed "%FILE8%="
   #sed "%FILE9%="
 close #sed

 'verify sed file existence before proceeding
    print "verifying SED file was created"
 #lablog, "verifying SED file was created"
 do
     res = fileExists(DestPath$,fnamenobas$;".sed")
  if res then exit do
   scan
 loop until res

'Check if iexpress.exe is installed (a built in Windows Install Maker = Self Extracting exe File)
 [makeinst]
 cursor hourglass
 print "@ - [makeinst] removing any existing ";DefaultDir$;"\EXE\";fnamenobas$;fixeddate$;".exe"
#lablog, "@ - [makeinst] removing any existing ";DefaultDir$;"\EXE\";fnamenobas$;fixeddate$;".exe"
     if fileExists(DefaultDir$;"\EXE", fnamenobas$;".exe") <> 0 then kill DefaultDir$;"\EXE\";fnamenobas$;".exe"
     if fileExists(DefaultDir$;"\EXE", fnamenobas$;fixeddate$;".exe") <> 0 then kill DefaultDir$;"\EXE\";fnamenobas$;fixeddate$;".exe"
     if fileExists(DefaultDir$, fnamenobas$;".exe") <> 0 then kill DefaultDir$;"\";fnamenobas$;".exe"
     print "@ [makeinst] exe file existence verified heading to execute IEXPRESS commandline"
 #lablog, "@ [makeinst] exe file existence verified heading to execute IEXPRESS commandline"

'makes 64 bit exe
    if bit=32 then [do32bit]
 print "running iexpress(64bit) commandline using the sed (information file) created earlier"
 #lablog, "running iexpress(64bit) commandline using the sed (information file) created earlier"
 'run iexpress commandline using the sed file created (sort of like an ini file)
 express64$ = "C:\Windows\System32"
 res=fileExists(express64$,"iexpress.exe")
   if res then run "iexpress /N /q ";sedfile$ else noiex=1 : goto [noiex]

'makes 32 bit exe
 [do32bit]
 print "running iexpress.exe(32bit) commandline using the sed (information file) created earlier"
 #lablog, "running iexpress.exe(32bit) commandline using the sed (information file) created earlier"
   if bit = 64 or bit = 0 then [verifyEXE]
   express32$ = "C:\Windows\SysWOW64"
 res=fileExists(express32$,"iexpress.exe")
   if res then run "iexpress /N /q ";sedfile$ else noiex=1 : goto [noiex]

call pause 5500

[verifyEXE]
print "@ [verifyEXE] entering verification loop"
 #lablog, "@ [verifyEXE] entering verification loop"
 #lablog, "DestPath$;\EXE\exe$ =  ";DestPath$;"\EXE\";exe$
'verify the exe file was created loop until it exists
 do
 res = fileExists(DestPath$;"\EXE",exe$)
  if res then exit do
 scan
  loop until res

'The EXE file gets created partially and fools the verification - pause to allow time
'for complete file creation - NOTE - This pause may need adjustment on YOUR PC
   call pause 4500

[noiex]
' copy SED script file to SED dir
  res = fileExists(DefaultDir$, fnamenobas$;".sed")
   if res and sed = 1 then
  print "@ - [noiex] - copying SED script file to SED dir if user chose to"
 #lablog, "@ - [noiex] - copying SED script file to SED dir if user chose to"
      open fnamenobas$;".sed" for input as #file
            open DestPath$;"\SED\";fnamenobas$;fixeddate$;fixedtime$;".sed" for output as #1
             print #1, input$(#file, lof(#file));
              close #file
               close #1
   end if

'keep autosave vbs script if user chose to
 res = fileExists(DefaultDir$, autoSave$)
   if res and sed = 1 then
      print "copying autoSave.vbs (if user chose to keep) to VBS dir"
      #lablog, "copying ";autoSave$;" to";DefaultDir$;"\VBS dir"
      open autoSave$ for input as #file
       open DestPath$;"\VBS\";autoSave$ for output as #1
        print #1, input$(#file, lof(#file));
        close #file
       close #1
   end if

 print DefaultDir$;"\EXE\";exe$;"  was created sucessfully"
 #lablog, DefaultDir$;"\EXE\";exe$;"  was created sucessfully"
  print "renaming EXE\ ";fnamenobas$;" to ";fnamenobas$;fixeddate$;".exe"
 #lablog, "renaming EXE\ ";fnamenobas$;" to ";fnamenobas$;fixeddate$;".exe"
  if fileExists(DefaultDir$;"\EXE", fnamenobas$;".exe") = 0 then
          notice fnamenobas$;".exe";" was Not created in  ";DefaultDir$;"\EXE"
  else
         if fileExists(DefaultDir$;"\EXE", fnamenobas$;".exe") <> 0 then
             name DefaultDir$;"\EXE\";fnamenobas$;".exe" as DefaultDir$;"\EXE\";fnamenobas$;fixeddate$;".exe"
              if tkn = 0 or tkn = 1 then
                    tkn = 3
                    print "sending to [newKey]/[continue] to add to ";categorie$;" List........"
                    #lablog, "sending to [newKey]/[continue] to add to ";categorie$;" List........"
                    newKey$ = fnamenobas$
                    categorie$ = programs$
                     goto [continue]
              end if
        end if
  end if

[done]
  print "@ - [done] - !SUCCESSFUL MISSION! NO ERRORS"
  #lablog, "@ - [done] - !SUCCESSFUL MISSION! NO ERRORS"
   yeserror = 0
 if spritecreated = 1 then run DefaultDir$;"\SpriteCreator v2\SpriteCreator.exe"
  spritecreated = 0
  cursor normal
 wait

'close bas2exe window
[quit.pick]
  #lablog,"@ - [quit.pick] closing #pick window "
 close #pick
wait

'close bas2exe window
[cancel]
  #lablog,"@ - [cancel] closing #pick window "
 close #pick
wait

sub cleanup
cursor hourglass
'delete tempBas.bas
       print "deleting tempBas.bas..............."
       #lablog, "deleting tempBas.bas..............."
    if fileExists(DefaultDir$,"tempBas.bas") <> 0 then kill DefaultDir$;"\";"tempBas.bas"

' Delete .vbs file temp .txt and temp .bas files
      print "Deleting temp .txt and temp .bas files"
     #lablog, "Deleting .vbs file temp .txt and temp .bas files"
       if fileExists(DefaultDir$,fnameTemp$) <> 0 then kill fnameTemp$
       if fileExists(DefaultDir$,"temp.txt") <> 0 then kill "temp.txt"

'delete folderdialog vbs script
      print "Deleting folderdialog vbs script"
     #lablog, "Deleting  folderdialog vbs script"
  if fileExists(DefaultDir$, "FolderDialog.vbs") then kill "FolderDialog.vbs"

'delete autosave.vbs
      print "Deleting autosave.vbs script"
     #lablog, "Deleting  autosave.vbs script file"
  if fileExists(DefaultDir$,"autoSave.vbs") <> 0 then kill DefaultDir$;"\";"autoSave.vbs"

'Deleting saved tkn file on user request
     print "Deleting saved tkn file on user request"
     #lablog, "Deleting saved tkn file on user request"
  if fileExists(DestPath$;"\TKN", fnamenobas$;fixeddate$;".tkn") <> 0 and tkn = 1 then kill DestPath$;"\TKN\";fnamenobas$;fixeddate$;".tkn"
  if res and tkn = 1 then kill DestPath$;"\TKN\";fnamenobas$;fixeddate$;".tkn"

'Deleting saved source code on user request
     print "Deleting saved tknfile on user request"
     #lablog, "Deleting saved tknfile on user request"
  if fileExists(DestPath$;"\BAS", fnamenobas$;fixeddate$;".bas") <> 0 and nobas = 1 then kill DestPath$;"\BAS\"; fnamenobas$;fixeddate$;".bas"

'delete the current project dir and files(copied bas file, tkn file, sll,dll, lbrun2.exe(renamed file)
'if user chose to "not include" this project dir
[remprogdir]
  print "deleting current project dir and files(copied bas file, tkn file, sll,dll, lbrun2.exe(renamed file) if user chose to not include this project dir"
 #lablog, "deleting current project dir and files(copied bas file, tkn file, sll,dll, lbrun2.exe(renamed file) if user chose to not include this project dir"
          if spritecreator = 1 then [done]
  if project = 0 then
          if fileExists(DestPath1$, "vbas31w.sll") <> 0 then kill DestPath1$;"\";"vbas31w.sll"
          if fileExists(DestPath1$, "voflr31w.sll") <> 0 then kill DestPath1$;"\";"voflr31w.sll"
          if fileExists(DestPath1$, "vthk31w.dll") <> 0 then kill DestPath1$;"\";"vthk31w.dll"
          if fileExists(DestPath1$, "vtk1631w.dll") <> 0 then kill DestPath1$;"\";"vtk1631w.dll"
          if fileExists(DestPath1$, "vtk3231w.dll") <> 0 then kill DestPath1$;"\";"vtk3231w.dll"
          if fileExists(DestPath1$, "vvm31w.dll") <> 0 then kill DestPath1$;"\";"vvm31w.dll"
          if fileExists(DestPath1$, "vgui31w.sll") <> 0 then kill DestPath1$;"\";"vgui31w.sll"
          if fileExists(DestPath1$, "vvmt31w.dll") <> 0 then kill DestPath1$;"\";"vvmt31w.dll"
          if fileExists(DestPath1$, fnamenobas$;".exe") <> 0 then kill DestPath1$;"\";fnamenobas$;".exe"
          if fileExists(DestPath1$, fnamenobas$;".tkn") <> 0 then kill DestPath1$;"\";fnamenobas$;".tkn"
          if fileExists(DestPath1$, fnamenobas$;".bas") <> 0 then kill DestPath1$;"\";fnamenobas$;".bas"

'delete the current project dir
          print "Deleting the current project dir"
          #lablog, "Deleting the current project dir"
        if pathExists(DestPath1$) then deldir = rmdir(DestPath1$)
  end if
'delete the root sed file
      print "deleting the root sed file..............."
    #lablog, "deleting the root sed file..............."
  if fileExists(DefaultDir$, fnamenobas$;".sed") <> 0 then kill DefaultDir$;"\";fnamenobas$;".sed"

'delete temp files left over from failed EXE creation
#lablog, "Deleting left over FAILED EXE temp files (if they exist)"
   if fileExists(DefaultDir$;"\EXE", "~";fnamenobas$;".CAB") <> 0 then kill DefaultDir$;"\EXE\~";fnamenobas$;".CAB"
   if fileExists(DefaultDir$;"\EXE", "~";fnamenobas$;".DDF") <> 0 then kill DefaultDir$;"\EXE\~";fnamenobas$;".DDF"
   if fileExists(DefaultDir$;"\EXE", "~";fnamenobas$;".CAB") <> 0 then kill DefaultDir$;"\EXE\~";fnamenobas$;".RPT"
   if fileExists(DefaultDir$;"\EXE", "~";fnamenobas$;"_LAYOUT.INF") <> 0 then kill DefaultDir$;"\EXE\~";fnamenobas$;"_LAYOUT.INF"
     cursor normal
 end sub

'function for checking file existence
 function fileExists(path$, filename$)
   #lablog,"@ - fileExists(path$, filename$) function"
  dim info$(0, 0)
  files path$, filename$, info$()
  fileExists = val(info$(0, 0)) 'non zero is true
   end function

'function for checking folder existence
 function pathExists(path$)
   #lablog,"@ - function pathExists(path$) function"
  pathExists = (mkdir(path$)=183)
   end function

'functions for making the folder dialog window
function FolderDialog$(caption$)
#lablog, "@ function FolderDialog$(caption)..........."
    WindowWidth = 600
    WindowHeight = 370
    UpperLeftX=INT((DisplayWidth-WindowWidth)/2)
    UpperLeftY=INT((DisplayHeight-WindowHeight)/2)
    BackgroundColor$ = "lightgray"
    ForegroundColor$ = "black"
    gosub [FolderDlgGetDrives]
    statictext #folderdlg.S, "Note: - Only Drives and Folders Appear Below - No Files Appear", 45, 15, 550, 25
    statictext #folderdlg.S, "Select a Drive or a Folder From the List", 175, 40, 300, 25
    statictext #folderdlg.D, "      (Double Click Drive Letters and Folders to Select or Navigate)", 85, 70, 395, 15
    listbox #folderdlg.list, FolderList$(, [FolderDlgSelect], 22, 90, 550, 130
    button #folderdlg.default, "Ok", [FolderDlgOk], UL, 190, 293, 85, 35
    button #folderdlg.B, "Back", [FolderDlgBack], UL, 490, 45, 80, 30
    button #folderdlg.C, "Cancel", [FolderDlgCancel], UL, 290, 293, 85, 35
    textbox #folderdlg.text, 42, 225, 510, 30
    statictext #folderdlg.path, "Selected Drive or Folder Path Appears Here", 130, 258, 400, 20
    open caption$ for dialog_modal as #folderdlg
     #folderdlg, "trapclose [FolderDlgCancel]"
     #folderdlg.default, "!font Arial 8 bold"
     #folderdlg, "font Arial 10 bold"
     #folderdlg.S, "!font Arial 10 bold"
     #folderdlg.path, "!font Arial 10 bold"
     #folderdlg.list, "font Arial 10 bold"
     #folderdlg.C, "!font Arial 10 bold"
     #folderdlg.D, "!font Arial 8 bold"
      #folderdlg.text, "!font Arial 10 bold"
 wait
[FolderDlgSelect]
    #folderdlg.list, "selection? temp$"
    if temp$ <> "" then
        level = level+1
        folder$ = folder$; temp$; "\"
        #folderdlg.text, folder$
         gosub [FolderDlgGetDir]
        #folderdlg.list, "reload"
    end if
 wait
[FolderDlgBack]
    if level > 0 then
        level = level-1
        if level = 0 then
            folder$ = ""
            gosub [FolderDlgGetDrives]
            else
            i = len(folder$)-1
            while mid$(folder$, i, 1) <> "\" and mid$(folder$, i, 1) <> ""
                i = i-1
            wend
            folder$ = left$(folder$, i)
            gosub [FolderDlgGetDir]
        end if
        #folderdlg.text, folder$
        #folderdlg.list, "reload"
    end if
 wait
[FolderDlgGetDrives]
    c = 1
    while word$(Drives$, c) <> ""
    c = c+1
    wend
    redim FolderList$(c)
    for i = 1 to c
    FolderList$(i) = word$(Drives$, i)
    next i
 return
[FolderDlgGetDir]
    files folder$, info$(
    s = val(info$(0,0))
    t = val(info$(0,1))
    redim FolderList$(t)
    for i = 1 to t
     FolderList$(i) = info$(i+s, 1)
    next i
 return
[FolderDlgOk]
    #folderdlg.text, "!contents? FolderDialog$"
    If right$(FolderDialog$,1) = "\" then FolderDialog$ = left$(FolderDialog$, len(FolderDialog$) - 1)
[FolderDlgCancel]
    close #folderdlg
 end function

'xxgeek code
sub fixdate
   print "@ - sub fixdate - fixing date ............"
   #lablog, "@ - sub fixdate - fixing date ............"
   fixDate$ = Date$("yyyy/mm/dd") 'set up the date format that works with a filename(remove the /)
   fix1$ = word$(fixDate$, 1, "/")
   fix2$ = word$(fixDate$, 2 ,"/")
   fix3$ = word$(fixDate$, 3 ,"/")
   fixeddate$ = "-";fix1$;"-";fix2$;"-";fix3$
 end sub

 sub fixtime
 print "@ - sub fixtime - fixing date ............"
   #lablog, "@ - sub fixtime - fixing date ............"
   fixTime$ = Time$() 'set up the time format that works with a filename(remove the /)
   fix1$ = word$(fixTime$, 1, ":")
   fix2$ = word$(fixTime$, 2 ,":")
   fix3$ = word$(fixTime$, 3 ,":")
   fixedtime$ = "-";fix1$;"-";fix2$;"-";fix3$
 end sub

'the following are cundo's lbsearch code edited by xxgeek to suit this app
'subroutine to search selection Help, and or Tutorial
 sub buttonClick h2$
 #lablog,"@ - sub buttonClick h2$"
   select case word$(h2$,2,".")
  case "search"
     #main.tb "!setfocus"
     #main.tb "!contents? searchFor$"
        searchFor$=trim$(searchFor$)
  if len(searchFor$)>2 then
            cursor hourglass
             redim searchList$(1000)
              for i = 1 to 1000 ' so so
                if helpList$(i)="" then
                  result$ = "yes"
                  #main.tb "!setfocus" : exit for
                end if
                  fileToOpen$=  word$(helpList$(i),2,chr$(0))
                   print helpFilePath$; "      "; fileToOpen$
                   open helpFilePath$; "\"; fileToOpen$ for input as #2
                    contents$ = input$(#2, lof(#2))
                    if instr(lower$(contents$), lower$(searchFor$)) then
                        count=count+1
                        searchList$(count)= helpList$(i)
                   end if
             close #2
         next i
              if count = 0 then prompt "No Entries Found for " + chr$(13) + searchFor$ + "       TRY AGAIN?" ; result$
                sort searchList$(), 0, count
              #main.listbox2 "reload"
              cursor normal
        else
            result$ = "yes"
            prompt "                3 Character Minimum"+chr$(13) +"                          TRY  AGAIN?";result$
        end if
     end select
 end sub
'cundo's  jbsearch code
'subroutine to open selected search item in a browser (htmlviewer if exists\default if not)
    sub lbDoubleClick h2$
    #lablog,"entering - sub lbDoubleClick h2$"
        #h2$ "selection? selection$"
     if selection$ = "" then exit sub
        fileToOpen$= word$( selection$,2,chr$(0))
        fileToOpen$=replace$( fileToOpen$ , "/", "\" )
    if fileExists(DefaultDir$, "htmlviewer.exe")  <> 0 then
           print "fileToOpen$ = ";helpFilePath$;"\";fileToOpen$ 'for testing with mainwin
           run "htmlviewer.exe ";helpFilePath$;"\";fileToOpen$
    else
           run "explorer.exe ";helpFilePath$;"\";fileToOpen$
    end if
 end sub

 function replace$( text$ , this$, tothis$ )
#lablog,"entering function replace$( text$ , this$, tothis$ )"
  while 1
        if instr(text$, this$) then
            f = instr(text$, this$)
            lenght=len(this$)
            text$ = mid$(text$,1,f-1);_
                tothis$;mid$(text$,f+lenght)
       else
             exit while
       end if
  wend
     replace$=text$
 end function

  sub combosub
 #lablog,"entering sub combosub"
 end sub

'sub to generate the window code and copy to clipboard, and texeditor
sub dummy fast$
 #lablog,"entering sub dummy fast$ to generate window code"
 print "entering sub dummy fast$ to generate window code"
  select case
       case fast$ = "#main.button1"
       #main.txt1 "!contents? txt$"
       #main.txt2 "!contents? theName$"
       #main.r1 "value? result$"
   if result$="set" then
     itag$="["
     otag$="]"
     closingCode$= "[quit]";chr$(13);_
                   " close ";txt$;chr$(13);_
                   " end"
    else
     closingCode$ = "Sub quit fast$";chr$(13);_
                    "  close #fast$"      ;chr$(13);_
                    "  end";chr$(13);_
                    "End Sub"
   end if
       #main.combo "selection? sel$"
    if instr(sel$,"popup") then
        includeButton$= "button ";txt$;".button1 ";chr$(34);_
         "&Exit";chr$(34);", "; itag$;"quit";otag$;", ul, 1, 1, 100, 30"
    end if
   toPrint$ = "WindowWidth = 640 : WindowHeight = 480";chr$(13);_
   "UpperLeftX=int((DisplayWidth-WindowWidth)/2)";chr$(13);_
   "UpperLeftY=int((DisplayHeight-WindowHeight)/2)";chr$(13);chr$(13);_
    includeButton$;chr$(13);chr$(13);_
    "Open ";chr$(34);theName$;chr$(34);" for ";sel$; " as ";txt$;chr$(13);_
    " ";txt$;" "; chr$(34); "trapclose ";itag$;"quit";otag$; chr$(34);chr$(13);_
    "wait";chr$(13);chr$(13);_
     closingCode$
     #main.ed "!cls"
     #main.ed toPrint$
     print "finished creating fastcode"
     #lablog, "finished creating fastcode"
     #main.ed "!selectall"
     #main.ed "!copy"
     #main.ed "!paste"
     #main.ed "!origin 0 0"
 end select
 end sub

'xxgeek's code
'quit program, save the current selected List first, and kill all htmlviewer
'windows if User chose to as well
[quit.main]
#lablog,"@ - [quit.main] calling saveValue, closing htmlviewers(if user chose to), calling cleanup, closing program"
   call saveValue
  print "closehtml -quit.main =  ";closehtml
 if closehtml <> 0 then run "taskkill /IM htmlviewer.exe   /F", HIDE
   call cleanup
 if lablogIsOpen = 1 then close #lablog
  close #main
 end

'sub to  create pauses in program
sub pause mil
#lablog,"entering sub pause mil"
    t=time$("ms")+mil
    while time$("ms")<t
        scan
    wend
end sub

'Carl Gundels Dictionary code (some editing by xxgeek to suit this app)
'sub to save current Dictionary Listings and text in texeditor
sub saveValue   'if the value is changed, save it
#lablog,"entering sub saveValue"
    if lastKey$ <> "" then
        #main.value "!modified? modified$";
        if modified$ = "true" then
            #main.value "!contents? saveThisValue$";
            call setValueByName lastKey$, saveThisValue$
            call collectGarbage
            call writeDictionary
        end if
    end if
end sub

'function to get selected Listing
function getKeys$(delimiter$)
#lablog,"entering function getKeys$"
  global keyCount
  pointer = 1
  while pointer <> 0
    'get the next key
        pointer = instr(dictionary$, chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134), pointer)
    if pointer then
      keyPointer = pointer + 9
      pointer = instr(dictionary$, chr$(134);chr$(165);chr$(134);"value";chr$(134);chr$(165);chr$(134), pointer)
      key$ = mid$(dictionary$, keyPointer, pointer - keyPointer)
      if instr(keyList$, chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134) + key$ + chr$(134);chr$(165);chr$(134);"value";chr$(134);chr$(165);chr$(134)) = 0 then
        getKeys$ = getKeys$ + key$ + delimiter$
        keyList$ = keyList$ + chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134) + key$
        keyCount = keyCount + 1
      end if
    end if
  wend
end function

'sub to write each Listing to corresponding file
sub writeDictionary
#lablog,"entering sub writeDictionary"
print categorie$
  open categorie$ for output as #writeDict
    #writeDict, dictionary$
  close #writeDict
end sub

'sub to read each Listing from corresponding file
sub readDictionary
#lablog,"entering sub readDictionary"
print "entering sub readDictionary"
  if fileExists(DefaultDir$, categorie$) <> 0 then
      open categorie$ for input as #readDict
    length = lof(#readDict)
    dictionary$ = input$(#readDict, length)
    close #readDict
    else
  end if
end sub

'sub to cleanup any mess in the dictionary text
sub collectGarbage
#lablog,"entering sub  collectGarbage"
  pointer = 1
  while pointer > 0
    'get the next key
    pointer = instr(dictionary$, chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134), pointer)
    if pointer then
      keyPointer = pointer + 9
      pointer = instr(dictionary$, chr$(134);chr$(165);chr$(134);"value";chr$(134);chr$(165);chr$(134), pointer)
      key$ = mid$(dictionary$, keyPointer, pointer - keyPointer)
      if instr(keyList$, chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134) + key$ + chr$(134);chr$(165);chr$(134);"value";chr$(134);chr$(165);chr$(134)) = 0 then
       value$ = getValue$(key$)
        newDictionary$ = chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134) + key$ + chr$(134);chr$(165);chr$(134);"value";chr$(134);chr$(165);chr$(134) + value$ + newDictionary$
        keyList$ = keyList$ + chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134) + key$ + chr$(134);chr$(165);chr$(134);"value";chr$(134);chr$(165);chr$(134)
      end if
    end if
  wend
  dictionary$ = newDictionary$
end sub

 sub setValueByName key$, value$
 #lablog,"entering sub  collectGarbage"
  dictionary$ = chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134);key$;chr$(134);chr$(165);chr$(134);"value";chr$(134);chr$(165);chr$(134)+value$+dictionary$
 end sub

'function to get info from selected Listing
 function getValue$(key$)
   print "Entering function getValue(key$......)"
#lablog, "Entering function getValue(key$......)"
  getValue$ = chr$(0)
  keyPosition = instr(dictionary$, chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134)+key$+chr$(134);chr$(165);chr$(134);"value";chr$(134);chr$(165);chr$(134))
  if keyPosition > 0 then
       keyPosition = keyPosition + 9  'skip over key tag
       valuePosition = instr(dictionary$, chr$(134);chr$(165);chr$(134);"value";chr$(134);chr$(165);chr$(134),  keyPosition)
      if valuePosition > 0 then
        valuePosition = valuePosition + 11   'skip over value tag
        endPosition = instr(dictionary$, chr$(134);chr$(165);chr$(134);"key";chr$(134);chr$(165);chr$(134), valuePosition)
       if endPosition > 0 then
        getValue$ = mid$(dictionary$, valuePosition, endPosition - valuePosition)
       else
        getValue$ = mid$(dictionary$, valuePosition)
       end if
    end if
  end if
end function

'sub to load selected categorie List
 sub loadKeys
#lablog, "Entering sub loadKeys........."
    keyList$ = getKeys$(chr$(134);chr$(165);chr$(134))
    redim keys$(keyCount)
    for item = 1 to keyCount
      keys$(item-1) = word$(keyList$, item, chr$(134);chr$(165);chr$(134))
    next item
    sort keys$(), 0 ,keyCount
  #main.keys "reload"
 end sub
'xxgeek code
'function to make custom messages
 function GetMessage$(message$)
#lablog, "Entering function GetMessage$(message$)......."
    WindowWidth = 520 :  WindowHeight = 740
    UpperLeftX=INT((DisplayWidth-WindowWidth)/2)
    UpperLeftY=INT((DisplayHeight-WindowHeight)/2)
    BackgroundColor$ = "lightgray" : ForegroundColor$ = "black"
        statictext #textmessage.text, "", 0, 0, 490, 600
        button #textmessage.button, "OK", [quit], UL, 230, 625, 35, 35
  open "Information" for window as #textmessage
       #textmessage.button, "!setfocus"
       print #textmessage, "trapclose [close]"
       #textmessage, "font Arial bold 10"
       #textmessage.text, message$
       #textmessage.button, "!font Arial bold 12"
  wait
[close]
      close #textmessage : exit function
      scan
  wait
[quit]
      scan
      close #textmessage : exit function
end function

'function to separate filename from full path to file
 function GetFilename$(fileName$)
#lablog, "Entering function GetFilename$(fileName$)......."
    i = len(fileName$)
    while mid$(fileName$, i, 1) <> "\" and mid$(fileName$, i, 1) <> ""
        i = i-1
    wend
    GetFilename$ = mid$(fileName$, i+1)
 end function

'function for making of popup menus (jb code from functions examples)
 function PopupMenu$(options$, width, bgColor$, textColor$, selBackColor$, selTextColor$)
#lablog, "Entering function PopupMenu$(etcetera......."
    'arguments:
    'options$ - comma-separated list of menu options
    'width - window-width, default = 100
    'bgColor$ - background color of the dialog
    'textColor$ - color of inactive text
    'selBackColor$ - backcolor of active, selected text
    'selTextColor$ - color of active, selected text
    'NOTE: colors are either a string of rgb values, one of the windows colours or
        'empty string (use default colour scheme)
    while word$(options$, count+1, ",") <> ""
        count = count+1
    wend
    height = count*20+38
    width = int(width) : if width < 100 then width = 100
    if bgColor$ = "" then bgColor$ = "white"
    if textColor$ = "" then textColor$ = "black"
    if selBackColor$ = "" then selBackColor$ = "darkblue"
    if selTextColor$ = "" then selTextColor$ = "white"
    WindowHeight = height
    WindowWidth = width
    UpperLeftX = MouseX
    UpperLeftY = MouseY
    graphicbox #popup.graph, 0, 0, width, height
    open title$ for dialog_modal_nf as #popup
    #popup, "trapclose [popupDlgCancel]"
    #popup, "font ms_sans_serif 16 9"
    #popup.graph, "down; fill "; bgColor$
    #popup.graph, "color "; textColor$; "; backcolor "; bgColor$
    for i = 1 to count
        #popup.graph, "place 4 "; i*20 - 2
        #popup.graph, "\"; word$(options$, i, ",")
    next i
    #popup.graph, "flush"
    #popup.graph, "when mouseMove [popupDlgMove]"
    #popup.graph, "when leftButtonDown [popupDlgSelect]"
    wait

[popupDlgMove]
    this = (MouseY-3)/20 : if this >= 0 then this = this + 1
    this = int(this)
    if this <> selection then
        #popup.graph, "backcolor "; bgColor$; "; color "; bgColor$
        #popup.graph, "place 2 "; selection*20 - 16; "; boxfilled "; width-12; " "; selection*20+2
        #popup.graph, "color "; textColor$
        #popup.graph, "place 4 "; selection*20 - 2
        #popup.graph, "\"; word$(options$, selection, ",")
        if this > 0 and this <= count then
            #popup.graph, "backcolor "; selBackColor$; "; color "; selBackColor$
            #popup.graph, "place 2 "; this*20 - 16; "; boxfilled "; width-12; " "; this*20+2
            #popup.graph, "color "; selTextColor$
            #popup.graph, "place 4 "; this*20 - 2
            #popup.graph, "\"; word$(options$, this, ",")
        end if
        selection = this
    end if
    wait
[popupDlgSelect]
    this = (MouseY-3)/20 : if this >= 0 then this = this + 1
    this = int(this)
    if this > count or this < 1 then wait
    PopupMenu$ = word$(options$, this, ",")
[popupDlgCancel]
    close #popup
    end function

'lb progress bar - Edited by xxgeek to suit this app
 sub progressBar
  cursor hourglass
#lablog, "Entering sub progressBar........."
[launch]
     WindowWidth = 500
     WindowHeight = 80
     UpperLeftX = INT((DisplayWidth-WindowWidth)/2)
     UpperLeftY = INT((DisplayHeight-WindowHeight)/2)+175
        graphicbox #progress.bar, 20, 15, 450, 10
        statictext #progress.text, "", 30, 25, 450, 25
     open "ProgressBar" for dialog_popup as #progress
      #progress, "trapclose [endprogress]"
      #progress.bar, "backcolor 100 100 250"
      #progress.bar, "down"
      #progress, "font arial 12 bold"
      #progress.bar, "setfocus"
      #progress.bar, "cls"
  while bar < 450
      #progress.bar, "boxfilled "; bar ;" 20"
         progress = bar/450
         progress =int(progress * 100)
       if pnum = 0 then #progress.text, "Getting User Path                ";progress; "  % done"
       if pnum = 1 then #progress.text, "Loading Dialogs List                ";progress; "  % done"
       if pnum = 2 then #progress.text, "Loading Reserved Words List                ";progress;  "% done"
       if pnum = 3 then #progress.text, "Loading Samples List               ";progress; "% done"
       if pnum = 4 then #progress.text, "Loading Functions List               ";progress;  "% done"
       if pnum = 5 then #progress.text, "Loading ASCII List              ";progress;"  % done"
   for break = 1 to 200
   next break
         bar = bar + 2.5
 wend
         bar = 0
 close #progress
  if pnum = 0 then
     WindowWidth = 120
      WindowHeight = 90
       UpperLeftX = INT((DisplayWidth-WindowWidth)/2)
        UpperLeftY = INT((DisplayHeight-WindowHeight)/2)+255
          statictext #ready.text, " Please Wait", 10, 20, 90, 20
       open "Ready" for dialog_popup as #ready
          #ready, "trapclose [endprogress]"
          #ready.text, "!font arial 10 bold"
          exit sub
  end if
       if pnum = 5 then
           #ready.text, "       Ready"
            for readymessage = 1 to 1000000
            next readymessage
          close #ready
      end if
 [endprogress]
  cursor normal
end sub

'xxgek's code
'get any errors that occur and let user know what's up and create a log file
[errorReport]
 #lablog, "@ - [errorReport]  -  calling sub cleanup  ........."
     call cleanup
     yeserror = 1
 #lablog, chr$(13);"Date ";date$();" ";Time$();" Error # ";Err;"  ";Err$;chr$(13)
    print chr$(13);"Date ";date$();" ";Time$();" Error # ";Err;"  ";Err$;chr$(13)
 #lablog, "Error # ";Err;"    ";Err$;chr$(13);chr$(13);" !  MISSION INTERUPTED  !  ";chr$(13);" !  ABORTING MISSION  ! ";chr$(13)
     print "Error # ";Err;"    ";Err$;chr$(13);chr$(13);" !  MISSION INTERUPTED  !  ";chr$(13);" !  ABORTING MISSION  ! ";chr$(13);" >>> Cleaning up Temp Files";chr$(13);chr$(13);" See lbHelpLabError.log "
    open "lbHelpLabError.log" for append as #1
     #1,  chr$(13);date$();" ";Time$();" Error # ";Err;"  ";Err$;chr$(13)
   close #1
 notice "Error # ";Err;"    ";Err$;chr$(13);chr$(13);" !  MISSION INTERUPTED  !  ";chr$(13);" !  ABORTTING MISSION  ! ";chr$(13);" >>> Cleaning up Temp Files...........";chr$(13);chr$(13);" See lbHelpLabError.log "




