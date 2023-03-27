# apex-copy-excel-to-ig

This Dynamic Action (DA) plug-in supports Copy - Paste, Import and Appending actions from an Excel file to an editable Interactive Grid (IG) report.

Minimum requirement: Oracle Application Expresss 19.1

This plug-in uses the <a href="https://sheetjs.com/" rel="nofollow">SheetJS</a> library CE edition.

<img width="950" alt="image" src="https://user-images.githubusercontent.com/100072414/227924485-0c9968d6-0f89-4161-8a45-4eac47abaa2a.png">

# setup

You can check each setup in my downloadable <a href="https://github.com/baldogiRichard/plug-in-site" rel="nofollow">Sample Application: APEX Plug-ins by Richard Baldogi</a>

There are 3 Dynamic Action types which can be used:

<b>Import from file</b>
<p>In this case the user can use a file browser where the excel file will be imported from:</p>
<p>The import consist of 2 parts:</p>
<ul>
  <li>Prepare: This action will initialize some variables in order to execute the insert of the file to the Interactive Grid.</li>
  <li>Execute: The selected sheet will be imported to the IG. Note: A select list must be defined in the Select List attribute.</li>
</ul>
<br>
<b>Paste from clipboard</b>
<br>
<p>In this case the copied data from the excel file will be pasted to the IG. Please note that the insertion will work only if the user first copies the data (CTRL+C) then pastes it to the site (CTRL+V) and then clicks the button if it were created under a click event.</p>
<b>Export selected to Excel</b>
<br>
<p>There are 2 logics defined:</p>
<ul>
 <li>When a file is selected in the browser and record(s) are selected in the IG:<br> In this case the headers will be used from the file and the selected rows from the IG will be appended to the file.</li>
 <li>Only record(s) are selected from the IG: <br> In this case the headers and the selected rows will be exported with the name and extension that was defined in attribute "filename" and attribute "extension".</li>
</ul>
