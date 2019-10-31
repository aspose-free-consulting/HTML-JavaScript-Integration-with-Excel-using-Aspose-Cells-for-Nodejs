# HTML & JavaScript Integration with Excel using Aspose.Cells for Nodejs

 This Aspose free consulting project helps you understand HTML/JavaScript integration with Excel using Aspose.Cells for Nodejs via Java. This helps you calculate and move data between HTML page and Excel sheet without opening the Excel file.
 
 # System Requirements and Notes
 
 ## Set up Modules and Environment

We need to setup up environment and all the required modules with configurations i.e node.js, nod-java bridge and express server. We used (Aspose.Cells for Node.js via Java v19.8)[https://downloads.aspose.com/cells/nodejs].

* It is better to install node.js v11.5.0 as we found that latest version (v12.0) has some compatibility issues with VS.NET tools (which are also required), so we have to use this version (downloaded and installed using node-v11.15.0-x64.msi). I have VS.NET 2017 on my pc (OS: Windows 10).
* Open “Node.js command prompt” as an Administrator and run the following commands by changing to your specified directory. In our case, it is “F:\files\”:
'''
** mkdir aspose.cells.js.java
** cd aspose.cells.js.java
** npm config set msvs_version 2017
** npm install -g node-gyp
** npm install java
'''
* Make sure that you have installed Oracle JDK/JRE 8 on the system and set all the relevant environment variables (e.g. JAVA_HOME, Path, classpath, etc.)
* Download Aspose.Cells for Node.js via Jav v19.8 and extract it into "aspose.cells.js.java/node_modules" folder
* Now install express (http) application server
** npm install express –g
* Note, when we use "npm install express -g" and then running the command “node server.js" could throw exception as it could not find "express". So, we have to use the following command: 
* npm install express –save
* However, with save option "aspose.cells" component folder under “aspose.cells.js.java/node_modules" folder was deleted, I had to copy it again, so you may do the same if you find such issue

#  The Demo

* The demo includes two files (check the code segments and scripts in the files for reference):
** server.js: in this file, we have code segment with Aspose.Cells for Node.js via Java API which reads and writes values (using reqest/response architecture), from/into the included XLSX file 
** The included file is taken as template file, code will need to be adjusted based on the input Excel file
** We will use express application server to request and response data as JSON
* Aspose.html: this is a simple frond end file which has some interface fields. You will input values in respective fields and then this will be submitted to server (which is listening at some port), and then the server.js file would come in play  
* Note, in server.js, the request data would be read and inserted into respective fields in the sheet and then we calculate formulas using Workbook.calculateFormula() method, So other cells with formulas will be processed in the sheet. 
* Finally, it will respond with calculated results as JSON to the client end  
* In .html file the other respective fields would be filled from JSON (response) data of the server

# How to Run the Demo

* Run "node server.js" in cmd prompt (as admin)
* Open http://localhost:8080/Aspose.html
* Click "submit"

Make sure, you have already installed all the requied components and modules including "aspose.cells" and "express" server.


# Screenshots

*	Make sure you have copied aspose.cells folder

<img width="468" alt="image" src="https://user-images.githubusercontent.com/1214951/67923917-09a42600-fbd1-11e9-8473-7348958adc66.png">
 
* Copy the demo files with template file in the folder
 
 <img width="468" alt="image" src="https://user-images.githubusercontent.com/1214951/67923926-0f9a0700-fbd1-11e9-8499-e5b4b927d3b6.png">
 
* Run the server.js file first

<img width="468" alt="image" src="https://user-images.githubusercontent.com/1214951/67923931-158fe800-fbd1-11e9-93c3-64c7483d9ff5.png">
 
* Open some browser (I used Google chrome) and run the HTML file by application server at port 8080
(The values shown into these fields are default values shown which are mentioned in the source of the .html file.)
  
  <img width="468" alt="image" src="https://user-images.githubusercontent.com/1214951/67923944-1aed3280-fbd1-11e9-80c2-95e9a36e7c09.png">

* Now click “submit” and the below fields would be calculated accordingly

<img width="468" alt="image" src="https://user-images.githubusercontent.com/1214951/67923955-1fb1e680-fbd1-11e9-8d8a-346164baf278.png">

* After changing the top values a bit and then clicking “submit” again. You will see the formula values are re-calculated and the results are shown into respective fields accordingly

<img width="468" alt="image" src="https://user-images.githubusercontent.com/1214951/67923968-23456d80-fbd1-11e9-89ef-603a42dc8d1e.png">


 ## Interested in Aspose free consulting project?
[If you are also interested in a free consulting project by Aspose team then please view details on this page](https://aspose-free-consulting.github.io/)

If you have any questions about Aspose APIs, please feel free to [post your query in Aspose file format APIs Forums](https://forum.aspose.com/). 

Also, you can keep in touch with the latest developments in [file format APIs offered by Aspose at our Blog](https://blog.aspose.com/). 


