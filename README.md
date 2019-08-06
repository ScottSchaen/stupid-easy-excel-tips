# Scott’s Stupid Easy Excel Tips and Tricks
<kbd>Windows</kbd> & <kbd>Mac</kbd>  
### Configuring your Excel for a better, faster, stronger experience 
#### *in 5 stupid easy steps*

1. [Make Excel (and Windows) snappier by disabling animations](#1-make-excel-and-windows-snappier-by-disabling-animations) <kbd>Windows Only</kbd>
2. [Opening Excel should take you to a new spreadsheet, not the welcome splash screen](#2-opening-excel-should-take-you-to-a-new-spreadsheet-not-the-welcome-splash-screen) <kbd>Windows</kbd><kbd>Mac</kbd>  
3. [Make the Quick Access Toolbar useful](#3-make-the-quick-access-toolbar-useful) <kbd>Windows</kbd><kbd>Mac</kbd>  
4. [Use the Status Bar — Don’t bother with unnecessary formulas](#4-use-the-status-bar--dont-bother-with-unnecessary-formulas) <kbd>Windows</kbd><kbd>Mac</kbd>  
5. [Customize the Home Ribbon — Load it up with *only* useful functions](#5-customize-the-home-ribbon--load-it-up-with-only-useful-functions) <kbd>Windows</kbd><kbd>Mac</kbd>  

Bonus: [MACROS!!!](https://github.com/ScottSchaen/excel-vba-macros)

- - - -

## 1. Make Excel (and Windows) snappier by disabling animations
<kbd>Windows Only</kbd>  

Windows adds some subtle animations to Excel (and other applications). You may not be able to put your finger on it, but it makes everything feel slower. Disabling these animations instantly makes Excel feel snappier. 

#### a) Disable worksheet animations
Selecting cells and scrolling through cells are “animated”. This is slowing you down. You’ll never look back after disabling it.

Animations (jumping to bottom)| No Animations (jumping to bottom)
--- | ---
![Scrolling with animations turned on](/images/scroll_before.gif)[]() | ![Scrolling with animations turned off](/images/scroll_after.gif)[]()

**How to do it:** <kbd>WIN</kbd> → Search for `Adjust the appearance and performance of Windows`. Uncheck `Animate Controls and elements inside windows`.

#### b) Disable animations when minimizing and maximizing windows
This is less of an Excel improvement and an overall Windows improvement. Like above, it’s subtle, but disabling it will make your PC feel faster. Don't worry about _smoothness_, Windows will still be plenty smooth.

Animations | No Animations
--- | ---
![Minimizing and maximizing windows with animations turned on](/images/window_before.gif)[]() | ![Minimizing and maximizing windows with animations turned off](/images/window_after.gif)[]()

**How to do it:** <kbd>WIN</kbd> → Search for `Adjust the appearance and performance of Windows`. Uncheck `Animate windows when minimizing and maximizing`.

- - - -

## 2. Opening Excel should take you to a new spreadsheet, not the welcome splash screen
<kbd>Windows</kbd> and <kbd>Mac</kbd>  

This one may be personal preference, but chances are when you open Excel you want to get straight to business with a fresh spreadsheet. If you're frequently opening recent files, you may want to leave it as is. 

Welcome Start Screen | Straight-to-business Worksheet
--- | ---
![Open to Excel Splash](/images/welcome_splash.png)[]() | ![Open to Worksheet](/images/new_sheet.png)[]()

**How to do it:** 
* **Windows**: `File` → `Options` → Uncheck “Show the Start screen when this application starts”
* **Mac**: `Excel` → `Preferences` → `General` → Uncheck "Show Workbook Gallery when opening Excel"

- - - -

## 3. Make the Quick Access Toolbar useful.
<kbd>Windows</kbd> and <kbd>Mac</kbd>  

By default, you’re multiple clicks away from opening files, creating new files, or printing. That's not right! Add all of these features and more to the quick access toolbar. They will always be there for you. This also works in other Microsoft applications!

<img src="/images/quick_access.png" width="500px" />

**How to do it:** 
* **Windows**: Click the dropdown in the Quick Access toolbar (on right) or right click anywhere on the ribbon and click `Customize Quick Access Toolbar`
* **Mac**: Click the dropdown in the Quick Access toolbar (on right) or `Excel` → `Preferences` → `Ribbon & Toolbar` → `Quick Access Toolbar`

- - - -

## 4. Use the Status Bar — Don’t bother with unnecessary formulas
<kbd>Windows</kbd> and <kbd>Mac</kbd>  

This is one of the biggest timesavers I have for you if you didn't already know about it. The Status Bar on the bottom right of the Excel Window is the easiest and fastest way to summarize data you’ve selected. You can very quickly get a feel for a dataset by utilizing this alone. Enable all of the calculations and immediately you’ll know the min and max values in the selection, the sum, the average, a count of non-empty cells, and a count of numbers. You can select a whole column and you don't need to worry about column headers and empty cells throwing off your data.

<img src="/images/status_bar.png" width="600px" />

**How to do it:** Right click anywhere on the status bar. Enable everything you want.

- - - -

## 5. Customize the Home Ribbon — Load it up with *only* useful functions
<kbd>Windows</kbd> and <kbd>Mac</kbd>  

The *HOME* tab of your ribbon is the default tab and of course your most used tab. Get rid of stuff you don't use and fill it with stuff that you _do_ use. There's a small work around to get this right. 

What you have (default):
![Default Ribbon](/images/ribbon_default.png)[]()

What you could have:
![Better Ribbon](/images/ribbon_you.png)[]()

What I have (Macros!):
![Ribbon with Macros](/images/ribbon_me.png)[]()

**How to do it:**
1. **Windows:** Right click anywhere on the ribbon and select `Customize the Ribbon...`
1. **Mac:** `Excel` → `Preferences` → `Ribbon & Toolbar`
2. The catch: You have to remove entire groups from the `Home` ribbon, like the `Cells` group. You can't remove one command/button from a group currently.
3. Add a new group, i.e. `Useful Stuff` to the `Home` tab and add every command/button you like.

**Recommendations:**
* Get rid of the `Cells` group, it's easier (and more organic) to right click on a cell/column/row and click `Insert`
* I use the `Filter` button and `Sort` buttons a lot. By default they're all bundled under `Sort & Filter`, so I get rid of that group and add separate buttons for `AutoFilter`, `Sort Ascending`, `Sort Descending`. You probably don't click `Find` if you are comfortable with <kbd>ctrl+f</kbd>, so you can get rid of the `Editing` group and add the buttons you like back in.
* If you use keyboard shortcuts for cut/copy/paste, get rid of `Clipboard`, but be sure to add the `Format Painter` back in.
* Add `Freeze panes`, otherwise you have to navigate to the `View` tab for this very useful feature. If you don't know what this is, you probably want to be using this.
* Scroll through all the `Popular Commands` in the custromization tool and add anything useful. Don't worry if a command already exists in another tab. For instance `PivotTable` is a good one to have handy. 
* It's easy to `reset customizations` if you don't like it. In the customizer you can "Choose commands from: `Main Tabs`" to see what you're missing from the default. You can also `Export all customizations`

**For the Power Users:**
You can add macros as commands/buttons and they will look and feel native. You can do about anything with VBA macros. Check out my [Excel VBA Macros](https://github.com/ScottSchaen/excel-vba-macros) and a How-To in my [Github](https://github.com/ScottSchaen/excel-vba-macros)
  
- - - -

## Did I miss anything?
Yes, I definitely did. Propose a file change by forking this project or find me on [LinkedIn](https://www.linkedin.com/in/scottschaen/) and shoot me a note.

**Rock on,**  
**Scott**
