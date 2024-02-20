# Disarm Add-In Project 

## Setting Up Your Development Environment

#### Prerequisites
- Microsoft Word
- Node.js (the latest LTS version)


#### Run the Microsoft Word Add-in
###### Clone the repository
###### Navigate to the Disarm-Add-In directory
###### Install the latest version of the Yeoman Generator for Office Add-Ins 

```sh
  npm install -g yo generator-office
```

###### Build and run the application 
```sh
  npm start
```

It is possible that while running this  `
  npm start
` command for the first time, the Add-In will request admin access for starting the application.

If successful, a new Microsoft Window will appear, with the loaded Disarm Add-In in the right corner of the taskpane. Click on the icon to start it. To debug the Add-In, click the arrow in the top right corner and Attach Debugger. This will open a new window with the client-side developer tools attached.

## Code Architecture
The repository is composed of 3 folders:
- **Disarm-Add-In**: The Microsoft Add-In, written in JavaScript, HTML and CSS
- **Disarm-Server**: A minimal API .NET application written in C#
- **DisarmPythonResultGenerator**: A console application written in C#

### Disarm-Add-In
This application contains the majority of the code written for this project. It was created using the Microsoft-Add-In API, and the remainder of the code was added to 3 different files in the application. The user interface was written using only HTML and CSS, and the logic was coded using pure JavaScript. 


All written code is contained in the `/src/taskpane directory`, where all files are named taskpane, with their corresponding file extensions (.html, .css, .js).


The HTML and CSS files were written by following the design made in the existing DISARM project, with small changes being made because of the constraints created by the  Microsoft-Add-In API. One of the security constraints is that forms and pop-ups are not allowed. This is mediated by encapsulating form content into CSS classes where the visibility is changed based on the clicked button. 


`taskpane.js` contains all project logic, and is the most complex. The `Office.OnReady` function calls helper methods that are triggered by click events. These helper methods than call other methods. In total, there are about 20 methods, of varying complexities and lengths. 


As this is a client-side application, one of its limitations is that it cannot access local files. The existing Disarm project created a Microsoft Excel file for each user on its local machine, and used it to save added tags. In this project, that problem was mediated by scanning the tag's text and id from the Word file content. For saving the tags and their descriptions, two json files were created, with opposite hierarchies. These json files were served by the Disarm-Server application, which is explained in the next section.

### Disarm-Server
This application is a minimal API .NET application with 3 API endpoints:
- `GET /tags` - serves the json tags file
- `GET /techniques` - serves the json techniques file
- `POST /clauses` - runs a python script that extracts clauses from the sentence by running a console application on the server. The body of the request is as follows:

```sh
{
  "sentence": "string",
  "result": "string"
}
```

This application is currently deployed [here](https://disarm-test.housepilot.de/).



### DisarmPythonResultGenerator
This is a .NET console application that runs the python script on the server. The process is triggered by the Disarm-Server application. 
