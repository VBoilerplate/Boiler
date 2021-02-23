# VBA Boilerplate

## The idea 
Boilerplate is an Excel binary file with VBA code in it, which can be used for every new VBA project as a boilerplate.
Building a boilerplate, which is to be used by as a start point for every VBA project was long in my mind. Somewhere in 2016 I have decided to put all the useful VBA code that I am using in a single repository. The repository is  https://github.com/Vitosh/VBA_personal, and up to now it has more than 60+ :star: in GitHub and just 1 contributor except me. The reason for this is that it probably looks a bit unstructured and I am the only one who can somehow find his way among all these files. Anyway, this week I am having some free time, thus I have decided to restart the project again -  create an Excel binary file with VBA code in it, which can be used for every new VBA project as a boilerplate.

## The structure
On February 2020 I have decided to change the repository to the current one:
https://github.com/VBoilerplate/Boiler

## How can I use the boilerplate:
Simply download it and use it! Or go through the files in and check them. If you find something interesting, copy it to your project.

## Video tutorials:
[YouTube VBA Boilerplate Tutorials](https://www.youtube.com/playlist?list=PLHvb-qAb0DaE2WXKfOXXNNRkoW990S5lP)

## Where is the official documentation?
On the current document and here - [vitoshacademy.com/boilerplate](https://www.vitoshacademy.com/boilerplate/)

## What is inside the boilerplate:

<ul>
 	<li><strong>ConstantsAndPublic</strong>
<ul>
 	<li><em>The module provides a list of the used public constants in the whole project. Including one public variable, which is used to build the error report</em></li>
</ul>
</li>
 	<li><strong>ExcelAdditional</strong>
<ul>
 	<li><em>Various useful procedures are here. They somehow do not belong anywhere else so far:</em>
<ul>
 	<li>FreezeRow</li>
 	<li>UnfreezeRows</li>
 	<li>SumArray</li>
 	<li>ChangeCommas</li>
 	<li>BubbleSort</li>
 	<li>IsArrayAllocated</li>
 	<li>RangeIsZeroOrEmpty</li>
 	<li>MakeRandom</li>
 	<li>IsRangeHidden</li>
 	<li>ColumnNumberToLetter</li>
 	<li>IsValueInArray</li>
 	<li>Rgb2HtmlColor</li>
 	<li>NamedRangeExists</li>
 	<li>GetRgb</li>
 	<li>CopyValues</li>
 	<li>OnEnd</li>
 	<li>OnStart</li>
</ul>
</li>
</ul>
</li>
 	<li><strong>ExcelDates</strong>
<ul>
 	<li><em>Dates were always tough for Excel users. These were tested for quite a long time.</em>
<ul>
 	<li>GetLastDayOfMonth</li>
 	<li>GetFirstDayOfMonth</li>
 	<li>AddMonths</li>
 	<li>AddMonthsAndGetFirstDate</li>
 	<li>DateDiffInMonths</li>
</ul>
</li>
</ul>
</li>
 	<li><strong>ExcelFormatCell</strong>
<ul>
 	<li><em>Formatting a cell in Excel can be done in various ways. These are some quick ones:</em>
<ul>
 	<li>FormatAsDate</li>
 	<li>FormatAsPercent</li>
 	<li>FormatAsCurrency</li>
 	<li>FormatAsEurProM2</li>
 	<li>FormatRedAndBold</li>
 	<li>WhiteRows</li>
 	<li>WhiteCell</li>
 	<li>FormatFontColorToGrey</li>
</ul>
</li>
</ul>
</li>
 	<li><strong>ExcelLastThings</strong>
<ul>
 	<li><em>Last row, last column, etc... in Excel are a must, when you are working with VBA. Make sure that you are aware, that some of the code ignores hidden ranges:</em>
<ul>
 	<li>LastColumn</li>
 	<li>LastRow</li>
 	<li>LastUsedColumn</li>
 	<li>LastUsedRow</li>
 	<li>LocateValueRow</li>
 	<li>LocateValueCol</li>
 	<li>Increment</li>
 	<li>Decrement</li>
</ul>
</li>
</ul>
</li>
 	<li><strong>ExcelPrintToNotepad</strong>
<ul>
 	<li><em>Printing to a .txt file is a feature that everyone needs. The file is in <span class="lang:default decode:true crayon-inline ">ThisWorkbook.Path &amp; "\Info</span>  folder.</em>
<ul>
 	<li>PrintToNotepad</li>
 	<li>CodifyTime</li>
 	<li>MakeAllValues</li>
</ul>
</li>
</ul>
</li>
 	<li><strong>ExcelStructure</strong>
<ul>
 	<li><em>Changes in the structure of Excel are found here. Named ranges, printing PDFs, working with comments, styles, resetting and unlocking stuff is found here</em>
<ul>
 	<li>LockScroll</li>
 	<li>StyleKiller</li>
 	<li>DeleteName</li>
 	<li>CoverRange</li>
 	<li>PrintActiveSheetPDF</li>
 	<li>PrintPage</li>
 	<li>DeleteDrawingObjects</li>
 	<li>UnhideAll</li>
 	<li>UnprotectAll</li>
 	<li>HideNeededWorksheets</li>
 	<li>AddCommentToSelection</li>
 	<li>PrintArray</li>
 	<li>PrintAllNames</li>
 	<li>DeleteAllNames</li>
 	<li>DeleteCommentInSelection</li>
 	<li>SelectMeA1RangeEverywhere</li>
 	<li>HideShowComments</li>
 	<li>ResetAndUnlock</li>
 	<li>EnableMySaves</li>
 	<li>DisabledCombination</li>
 	<li>DisableShortcutsAndSaves</li>
</ul>
</li>
</ul>
</li>
 	<li><strong>ExcelVBE</strong>
<ul>
 	<li><em>Be <strong>careful</strong> here. In general, this one could be <strong>dangerous</strong>, as far as it has one sub named <span class="lang:default decode:true crayon-inline">ImportModules</span>. It imports all the modules from a given folder to a given workbook. The "problem" is that before importing these, it deletes all other modules there. Just make sure that you know what you are doing, before using any of the subs from there.</em>
<ul>
 	<li>PrintAllCode</li>
 	<li>PrintAllContainers</li>
 	<li>ListProcedures</li>
 	<li>ExportModules</li>
 	<li>GetFolderOnDesktopPath</li>
 	<li>CreateFolderOnDesktop</li>
 	<li>ImportModules</li>
 	<li>DeleteAllVba</li>
</ul>
</li>
</ul>
</li>
 	<li><strong>FormExample</strong></li>
 	<li><strong>FormSummaryPresenter</strong></li>
 	<li><strong>FrmExample</strong></li>
 	<li><strong>FrmInfo</strong>
<ul>
 	<li>The above four a combined together.  To run the form, call "ShowMainForm". It does the rest. The forms are built, as in the article here - <a href="https://www.vitoshacademy.com/vba-the-perfect-userform-in-vba/">the perfect userform</a></li>
</ul>
</li>
 	<li><strong>tblInput (Input)</strong>
<ul>
 	<li>There is 1 sub for selection_change in this one. It checks the Zoom.</li>
</ul>
</li>
 	<li><strong>tblSettings (Settings)</strong>
<ul>
 	<li>Nothing in this one. It is by default <span class="lang:default decode:true crayon-inline">xlVeryHidden</span><em><strong>. </strong></em>Its idea is to put some data inside, avoiding the data in <strong><em>ConstantsAndPublic</em>.</strong></li>
</ul>
</li>
 	<li><strong>TddMain</strong></li>
 	<li><strong>TddSpecDefinition</strong></li>
 	<li><strong>TddSpecExpectation</strong></li>
 	<li><strong>TddSpecInlineRunner</strong></li>
 	<li><strong>TddSpecSuite</strong>
<ul>
 	<li>The 5 modules and classes above are a framework taken from <a href="https://github.com/VBA-tools/vba-test">here</a>, with some small changes. <em><strong>TddMain</strong></em> is where the tests are.</li>
</ul>
</li>
 	<li><strong>VersionsAbout</strong>
<ul>
 	<li>Well, this is #VBA. I have seen lots of projects, where the versioning is inside, hidden in a module. This is probably not a good practice (again!). But so these stay there.</li>
</ul>
</li>
 	<li><strong>xl_main</strong>
<ul>
 	<li>Workbook_BeforeClose</li>
 	<li>Workbook_BeforeSave</li>
 	<li>Workbook_NewSheet</li>
 	<li>Workbook_Open</li>
</ul>
</li>
</ul>

:cactus::cat::dog::monkey:
