# HTE Platform

# Getting the Code to run on your Copy of the Platform

In order for the Platform to work, you need to amend /src/global Variables.js with the folder/file IDs of your particular situation. 

The globalVariableDict dictionary contains a key-value pair for each lab that is using the platform with the key being the file ID of the instance of the HTE Platform running the code and the value being a dictionary of all the file/folder IDs needed for operation. 

That way, one code-platform can be used to manage several instances of the platform for different labs (or a development version). In order to differentiate the functionality for different labs, the file IDs of the Google Sheets of the different labs act as switches to determine to which folder a file should be written or which template to be used. 

# Setting up from scratch using and working on VS Code

One option to work on this project is to use VS Code, which allows to perform all the git pull, commit and pull from a visual interface. The amount of command line directives is limited and is necessary only once to set up VS Code correctly. After that you can manage your code from the software itself. We use Windows PowerShell for all the steps that need to be completed from command line, as it has a nice interface.
Let's see how to set this up.

One can also use the built-in IDE of Google Apps Script, but if you want to make bigger changes or want to add other developers, it's highly advised to use git in combination with clasp to manage the code. 

## Software installation
Follow the steps in [this page](https://medium.com/geekculture/how-to-write-google-apps-script-code-locally-in-vs-code-and-deploy-it-with-clasp-9a4273e2d018) to install VS Code, Node.js, npm and clasp. 
You can leave the default options for most of the cases. If you have SSH already in place you might want to select the option which allows to use existing
SSH and not to download and create a new installation.
It might be that you need to close and reopen VS Code to see the google autocompletion activated.

## GitLab setup
Go to GitLab and login. In the list of your project, or in the Project information > Members tab, check that you have developer status.
If not, ask to the project creator to grant you this.

If new to GitLab, you will be prompted to set a new SSHkey and an identification token to be able to push/pull through HTTPS. Set them up as indicated in the links:
SSH key needs to be set up through command line (in PowerShell), and the public key manually copied in your GitLab user settings. The token name is public, 
so should be fairly generic and all the boxes can be ticked. Once generated, SAVE it (as suggested), as it will not be possible to visualise it again.

Alternatively, you can also create a Personal Access Token in Gitlab (pick an expiry date further away in the future, so you don't have to renew it soon and check the boxes with the rights that your code requires - I checked them all) and use this to add your account to the Gitlab extension in VSCode (Ctrl+Shift+p, Write Gitlab: Add account to VScode. It'll ask you for the domain and then for the token you just generated. Et Voila...)

## Get the code on VS Code
Now create on your PC a landing folder for the project code, ideally an empty folder.

In VScode, in the initial GetStarted page which opens in the editor when you open the software, you see a "Clone git repository" option. Click on it and copy paste there the HTTPS url of the Project.
You can find this going in the Project page on GitLab and clicking on the blue button "Clone" and copying the link for "Clone with HTTPS".
Back in VS Code you need to select a landing folder, the one you previously created.
It can be that a pop up appears for identification, with the GitLab logo. Insert here your username (same as Roche username), and the TOKEN that you obtained early on in setting up GitLab.

## Start working on the code
After a few seconds the code is now in you local folder. On the bottom left you see the branch in which you are working (e.g. main or development).
By clicking on it a menu appears and you can select the appropriate branch to work on. You can see local branches (with the branch symbol) and remote ones (with the cloud symbol). At the beginning you are checkout on the main branch by default. Hence, when clicking on it, you will see in the menu a local main branch (plus the remote ones, including the remote main). 
To switch to development for the first time, select the remote development branch. After this, a local copy of the development branch will be created, and you can work on there.


## Commiting changes for the first time
After you've made your first coding steps and want to commit the changes, go to the Source Control (Ctrl + Shift + G) ribbon and enter a message regarding this commitment. After clicking on the blue Commit Button, you will get an error message saying you need to define both username and email for changes.
Open a terminal in VS Code and paste both lines separately, changing the placeholders accordingly.
git config --global user.name "Your Name"
git config --global user.email "youremail@yourdomain.com" 
After submitting both statements, you can now commit changes to the project.



***
