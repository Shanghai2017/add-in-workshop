These steps to generate the base code for the add-in should be completed in lesson 1. If not, then do them now.
Also, node and git should already be installed per the prerequisites. 

2.1. go here:

<https://dev.office.com/getting-started/addins>

2.2. select "Excel"

2.3. scroll down and select "Other tools" (big, clickable image in the "Build | What development tool do you use?" box)

2.4. while following the steps to generate the add-in code using Yo, use the following options:

```
? Would you like to create a new subfolder for your project? Yes                                               
? What do you want to name your add-in? lesson2_code_ts                                                        
? Which Office client application would you like to support? Excel                                             
? Would you like to create a new add-in? Yes, I need to create a new web app and manifest file for my add-in.  
? Would you like to use TypeScript? Yes                                                                        
? Choose a framework: Jquery                                                                                   
                                                                                                               
For more information and resources on your next steps, we have created a resource.html file in your project.   
? Would you like to open it now while we finish creating your project? No                                      
```

Navigate to the src folder.

2.5. Edit index.html, and change the "Run" button text in this section to "Setup". It
should be on about line 53.

```html
        <!-- change id="run" to id="setup" -->
        <button id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
            <!-- change >Run< to >Setup< -->
            <span class="ms-Button-label">Run</span>
            <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--ChevronRight"></i></span>
        </button>
```

2.6. In app.ts, find this line and change all occurences on the line from "run" to "setup":

```typescript
      $('#run').click(run);
```

      like this: 

```typescript
      $('#setup').click(setup);
```

2.7. Also find this line and change the name of the function to "setup":

```typescript
      async function run() {
```        

      like this: 

```typescript
      async function setup() {
```

2.8. Now replace the following stock comment in function setup() from: 

```typescript
      /**
       * Insert your Excel code here
       */
```

      to: 

```typescript
    await Excel.run(async function (context) {

      try {

        var wSheetName = 'Sample';
        var sheet = context.workbook.worksheets.add(wSheetName);
        sheet.load('name');
        await context.sync();
        console.log(sheet.name);

        const data = [
          ["Product", "Qty", "Unit Price", "Total Price"],
          ["Almonds", 2, 7.50, "=C3 * D3"],
          ["Coffee", 1, 34.50, "=C4 * D4"],
          ["Chocolate", 5, 9.56, "=C5 * D5"]
        ];

        const range = sheet.getRange("B2:E5");
        range.values = data;
        const header = range.getRow(0);
        header.format.fill.color = "#4472C4";
        header.format.font.color = "white";

        sheet.activate();
        console.log("Added setup table to Sample sheet");
        return await context.sync();
      }
      catch (error) {
        console.log("There was an error in Excel.run()!");
      }
    });
```

2.9. Now that you have the setup code for lesson 2, we can add code to use Excel
functions. 

Now add a new button and handler to create a Grand Total under the Total Price column. For this, use the sum formula.

2.10 Add another button with a label of "Grand Total".

2.11 Add code to total the Total Price column and put the result in E7 (below the last entry). Also add the label "Grand Total" in B7.

This should be the result Grand Total Table
Hints:

- Use Excel Worksheets Functions
- Remember that the values array will index from 0, even though the Excel addresses are 1 based.
- Use the workbook.functions.sum() method.
- Notice that the sum() method returns programmatically the value of the sum of the range, which we add to a cell. However, typically, you'd add a formula like this: =sum(<range>) into that cell instead of the resulting value.

Note: if you succeeded in creating the Grand Total button and handler, go to step 2.7, otherwise continue.
Let's add the "Grand Total" button so we can total up the prices. In
index.html, add this next to the existing "Setup" button:
        
      <button id="grand-total" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
          <span class="ms-Button-label">Grand Total</span>
          <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--ChevronRight"></i></span>
      </button>

2.12. Add the handler code for the totaling button in app.js:

```javascript
      $("#grand-total").click(grandTotal);
```

2.13. And the grantTotal function:

```javascript
      function grandTotal() {
        Excel.run(async (context) => {
          var range = context.workbook.worksheets.getItem("Sample").getRange("E3:E5");
          var rangeTot = context.workbook.worksheets.getItem("Sample").getRange("B7:E8");
          var gTot = context.workbook.functions.sum(range);

          range.load("values");
          rangeTot.load("values");
          gTot.load();

          context.sync()
            .then(function () {
              console.log("Loaded values, adding =sum()");
              var vTot = rangeTot.values;

              console.log(gTot.value);
              console.log(range);
              vTot[0][3] = gTot.value;
              vTot[0][0] = "Grand Total";
              vTot[0][1] = "=sum(c3:c5)";

              rangeTot.values = vTot;

              return context.sync()
                .then(function () {
                  console.log("Added =sum() function");
                });

            });

        })
          .catch(function (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
          });
      }
```

2.14. You will be using the range calculation API in Lesson 4, so let's add another value into the Grand Total row. This one should total up the Qty column but not use the workbook.functions.sum() method. Instead add the =sum() formula into the cell for later calculation.

This should be the result Grand Total Table

2.15. After this is successful, add another row with Tax (say B6:E6) and include that into the Grand Total amount.      

