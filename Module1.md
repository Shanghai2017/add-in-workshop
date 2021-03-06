# Module 1: Warm Up

We will use the [Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/application). For information about the API requirements see [Excel JavaScript API Requirements](https://github.com/OfficeDev/office-js-docs/blob/master/reference/requirement-sets/excel-api-requirement-sets.md).

This page contains instructions for the Basic track in the first module in the workshop. The steps for the Advanced track are in [Module1_advanced.md](Module1_advanced.md).

The next module is [Module 2](Module2.md).

# Contents
* [Setup](#setup): 1.1 - 1.5
* [Run the First Sample](#run-the-first-sample): 1.6 - 1.10
* [Add Code](#add-code): 1.10.1 - 1.10.4
* [Run, Test, and Save](#run-test-and-save): 1.11 - 1.13

## Setup
Configure the *Script Lab* add-in in Excel.

- 1.1 Open the desktop version of Excel.

- 1.2 Go to **Insert** > **My Add-ins** > **Office Add-ins**.  
    ![alt text](Module_1_Office_Add-ins_button.png "Office Add-ins button")

- 1.3 Search for *Script Lab*.  
    ![alt text](Module_1_Script_Lab_in_store.png "Script Lab add-in in the Office Add-ins Store")

- 1.4 Click **Add**.

- 1.5 Click on the **Script Lab** ribbon.  
    ![alt text](Module_1_Script_Lab_ribbon.png "Script Lab ribbon")

## Run the First Sample
Use sample code to highlight a range of cells.

- 1.6 In the **Code** pane, select **Samples**.

- 1.7 Select **Basic API call (JavaScript)**.

- 1.8 In the **Run** pane, select **Basic API call (JavaScript)**.

- 1.9 Select a matrix of several cells (see screenshot below) and click the **Run Code** button.  
    For example, if four cells are selected  
    ![alt text](Module_1_cells_selected.png "Selected cells")  
    Observations:  
    The cells selected should be highlighted in yellow.  
    ![alt text](Module_1_cells_highlighted.PNG "Highlighted cells")  
    Review the code in the Code pane.  
    Notice the ```Excel.run()``` invocation  
    ```javascript

    $("#run").click(run);

    function run() {
        Excel.run(function (context) {
            var range = context.workbook.getSelectedRange();
            range.format.fill.color = "yellow";
            range.load("address");
            return context.sync()
                .then(function() {
                    console.log("The range address was \"" + range.address + "\".");
                });
        })
            .catch(function(error) {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            });
    }
    ```  
    Note the ```context.sync()``` and ```.then``` pattern. Asynchronous code must always return a Promise.  

    Notice the output from the ```console.log()``` call on line 10 on the Firebug Console tab.

Adding some functionality:

You can now edit the *Basic API call (JavaScript)* sample code and it will save it to your snippets.

- <a name="1.10" />1.10 Using the Script Lab code editor, modify the code to populate the cells with increasing numbers starting at 1.  
    For example, if you select four cells, the output will look like this:  
    ![alt text](Module_1_cells_highlighted_with_data.png "Highlighted cells with values")

**Hints:**

* Use the `.values` property of the Range object.
* Remember to load "values" first and then sync.
* Fewer calls to `context.sync` mean fewer calls to the Office application.
* If you use another function to populate, remember to pass in the context as well as the range.
* If you get stuck, look at the other sample snippets for ideas.
* For more information about Range objects click on the link below:  
    * [Range Object - JavaScript API for Excel](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/range.md)

**Note:** If you've already done this module on your own, congratulations! You can skip to step [1.12](#1.12). Otherwise, keep following the next steps.

### Add Code
In this section you will add code to the sample to populate the highlighted range with data.
- 1.10.1 Change the following line: 
    ```javascript
    range.load("address");
    ```  
    To: 
    ```javascript
    range.load([ "address", "values"]);
    ```

- 1.10.2 After the following line:  
    ```javascript
    console.log("The range address was \"" + range.address + "\".");
    ```  
    Add a call to a separate function in which we will populate the cells:  
    ```javascript  
    return populateRange(context, range); // Added this line of code  
    ```

- 1.10.3 Now add the following populate function after the run() function:  
    ```javascript
    // Added the following code block
    function populateRange(context: Excel.RequestContext, range: Excel.Range) {
        console.log("populateRange: range is - ", range.address);  
            var newValues = range.values;
            var counter = 1;  
            for (var i = 0; i < newValues.length; i++) {  
                for (var j = 0; j < newValues[i].length; j++) {  
                    newValues[i][j] = counter++;  
                }  
            }  
            range.values = newValues;  

            return context.sync()
	    .then(function () {
			    console.log("finished populating the matrix");
			    });
    }  
    ```  

- 1.10.4 Now run and test the code making sure that it works and shows as the image in [1.10](#1.10).

## Run, Test, and Save
Run the modified code, ensure it works, and save your code.

- 1.11 Once satisfied, run it to show the populated cells.

- <a name="1.12" />1.12 Now you have a modified version of the *Basic API call (JavaScript)* sample code in **My Snippets**. Navigate to see that it's there.  
    * To get **My Snippets**, click on the hamburger menu in the top left corner of the **Code** panel. Then click **My Snippets** at the top. You should see your snippet in the **My snippets on this computer** section.
    * This modified code in **My Snippets** will only remain in the add-in memory until you clear your browser cache.

- 1.13 Save your snippets on GitHub
    * If you want to save your snippet, you can do that by first signing in with GitHub on the **My Snippets** page.
    * Click on the snippet that you want to save. This will take you to the text editor portion of the **Code** panel.
    * Hover over the **Share** button and choose either **Public gist** or **Private gist**.
