# Limit Up Word-Add-in
Project Details
 # Title
    Authentication & UI Completion

 # Description
    The Word Add-in enables secure user authentication via Firebase and provides a user-friendly interface for seamless interaction. It also integrates AI-powered features using Llama 3 (8B) via Groq API, enhancing document editing and research capabilities.

## Initialize 
1. Clone and open the github project in VS Code if you have a repository link
2. Open terminal and install node_modeules if not using "npm install" command
3. Run "npm start" to start the project 


## how to run the debugger on mac
https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-office-add-ins-on-ipad-and-mac

# TLDR:
1. run this command:
  
  defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true

2. now right click, inspect

## Before Commit

to make sure everything works:

* npm run build
* npm run lint
