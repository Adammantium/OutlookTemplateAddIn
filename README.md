# OutlookTemplateAddIn
A simple Outlook AddIn that makes it easy to create custom ribbons and categories.  
------------------------
[![forthebadge](https://forthebadge.com/images/featured/featured-built-with-love.svg)](https://forthebadge.com)
[![forthebadge](https://forthebadge.com/images/featured/featured-powered-by-electricity.svg)](https://forthebadge.com)
------------------------

This plugin allows to add files and links to your menu, without any complicated programming or configuration. Just plain and simple.
To get Files to appear in your Menu, you just put them into a specific folder.

# How to setup
1. Install the AddIn
2. Create a folder in `%appdata%` with the name `OutlookTemplates`
3. Any folder directly below that will create a new `Tab` with the same name
4. Any folder inside a `Tab` folder will create a `category` with the same name
5. Put the files you want in your menu inside a `category` folder

### Supported Filetypes:
`.msg` `.oft` `.txt` `.html` `.lnk`  `.url`  
Depending on the filetype a icon will be chosen.  

### Custom Icons
If you want a file to have a custom icon, just add a .png file with the same name as the original file. So if you have `My Template.msg` the icon needs to be named `My Template.msg.png` for it to appear.


## Example
### Folder structure
![image](https://github.com/Adammantium/OutlookTemplateAddIn/assets/38858318/641af799-3177-4923-8c22-6e59de78dd2f)  
A folder structure like this will result in two tabs named `Dev` and `Online`.  
`Dev` will have one category called `Mailcow` with two clickable buttons. Those will open the default webbrowser with their containing links. The WebMail url will have a custom icon as well.
`Online` will have two categories called `Entertainment` and `Search Engines`.

### Result in Outlook
![image](https://github.com/Adammantium/OutlookTemplateAddIn/assets/38858318/55bf1c8b-5c11-43bf-b4a1-cf2c5bd501c0)  
![image](https://github.com/Adammantium/OutlookTemplateAddIn/assets/38858318/9eca7470-3458-46a4-b7d9-316b4aa391cd)


## Is it compatible with GPO?
Yes, just copy the folder structure to your users home folder and it will automatically load all the required data.  
Because its in the AppData folder, it will make sure that users will have the templates available even when offline or not connected to a central storage server.
### Silent Install
`"C:\Program Files\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe" /install OutlookTemplateAddIn.vsto /silent`
