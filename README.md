# Scott’s Stupid Easy Excel Configuration
### <kbd>Windows Edition</kbd> 
### Make Excel better, faster, stronger. Set it up for success.

## Make Excel (and Windows) snappier by disabling animations
Windows adds some subtle animations to Excel (and other applications). You may not be able to put your finger on it, but it makes everything feel slower. Disabling these animations instantly makes Excel feel snappier. 

#### Disable worksheet animations
Selecting cells and scrolling through cells are “animated”. You may not realize, but this feature is slowing you down. You’ll never look back after disabling it.

Animations (jumping to bottom)| No Animations (jumping to bottom)
--- | ---
![Scrolling with animations turned on](/images/scroll_before.gif)[]() | ![Scrolling with animations turned off](/images/scroll_after.gif)[]()

<kbd>WIN</kbd> → Search for `Adjust the appearance and performance of Windows`. Uncheck `Animate Controls and elements inside windows`.

#### Disable animations when minimizing and maximizing windows
This is less of an Excel improvement and an overall Windows improvement. Like above, it’s subtle and will make your computer feel faster without it. Don't worry about _smoothness_, Windows will still be plenty smooth.

Animations | No Animations
--- | ---
![Minimizing and maximizing windows with animations turned on](/images/window_before.gif)[]() | ![Minimizing and maximizing windows with animations turned off](/images/window_after.gif)[]()

<kbd>WIN</kbd> → Search for `Adjust the appearance and performance of Windows`. Uncheck `Animate windows when minimizing and maximizing`.

## Opening Excel should take you to a new spreadsheet, not the welcome splash screen
It’s personal preference, but chances are when you open Excel you want to get straight to business with a fresh spreadsheet. If you're frequently opening recent files, you may want to leave it as is. 

Welcome Start Screen | Straight-to-business Worksheet
--- | ---
![Open to Excel Splash](/images/welcome_splash.png)[]() | ![Open to Worksheet](/images/new_sheet.png)[]()

File → Options → Uncheck “Show the Start screen when this application starts”


## Make the Quick Access Toolbar useful.
By default, you’re multiple clicks away from opening files, creating new files, or printing. That's not right! Add all of these features and more to the quick access toolbar, which will always be there for you.

<img src="/images/quick_access.png" width="500px" />

Hit the dropdown in the Quick Access toolbar or right click anywhere on the ribbon and click `Customize Quick Access Toolbar`

## Use the Status Bar — Don’t bother with unnecessary formulas
This is one of the biggest tips I have for you if you didn't know about it. The Status Bar on the bottom right of the Excel Window is the easiest and fastest way to summarize data you’ve selected. Enable all of the calculations and immediately you’ll know the min and max values in the selection, the sum, the average, a count of non-empty cells, and a count of numbers. You can select a whole column and you don't need to worry about column headers and empty cells throwing off your data.

<img src="/images/status_bar.png" width="600px" />

Right click anywhere on the status bar. Enable everything you want.

## Customize the Home Ribbon — Load it up with only useful functions
The *HOME* tab of your ribbon is the default tab and the most used, of course. Get rid of stuff you don't use and fill it with stuff that you _do_ use. There's a small work around to make this work. 

What you have (default):
![Default Ribbon](/images/ribbon_default.png)[]()

What you could have:
![Better Ribbon](/images/ribbon_you.png)[]()

What I have (Macros!):
![Ribbon with Macros](/images/ribbon_me.png)[]()

**How to do it:**
1. Right click anywhere on the ribbon and select `Customize the Ribbon...`
2. You have to remove entire groups from the `Home` ribbon, like the `Cells` group. You can't remove one button from a group currently.
3. Add a new group, i.e. `Useful Stuff` to the `Home` tab and add every button you like.

**Recommendations:**
* Get rid of the `Cells` group, it's easier (and more organic) to right click on a cell/column/row and click `Insert`
* I use the `Filter` button and `Sort` buttons a lot. By default they're all bundled under `Sort & Filter`, so I get rid of that group and add separate buttons for `AutoFilter`, `Sort Ascending`, `Sort Descending`. Most people use <kbd>ctrl+f</kbd> to `Find`, and the other `Editing` buttons can be added back in.
* If you use keyboard shortcuts for cut/copy/paste, get rid of `Clipboard`, but be sure to add the `Format Painter` back in.
* Add `Freeze panes`, otherwise you have to navigate to the `View` tab for this very useful feature.
* Check out `Populate Commands`. Don't worry if a command already exists in another tab. For instance `PivotTable` is a good one to have handy. 
* It's easy to `reset customizations` if you don't like it. In the customizer you can "Choose commands from: `Main Tabs`" to see what you're missing from the default. You can also `Export all customizations`
