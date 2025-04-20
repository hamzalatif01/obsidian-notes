Go to **File** > **Save As**.


1. In the **"Save as type"** dropdown, select: **“Word Macro-Enabled Document (*.docm)”**

2. Save the file.

#### Insert a Button and Assign the Macro

This is the clickable "**Copy**" button.

How to do it:

1. Go to the **Developer tab** in Word.

If you don’t see the Developer tab:
Go to File > Options > Customize Ribbon → check the Developer box → click OK.


2. Click on your document to place the **button**.

3. In the Developer tab and under **ActiveX Controls**, choose the Button (**Command Button**).

4. Right-click the button and choose **Properties** to rename it or change the caption (e.g., make it say “Copy”).


5. Close the Properties box.


6. Right-click the button again and choose **View Code**.

7. Paste this **code**:
> Private Sub CommandButton1_Click()
> Dim rng As Range
> Set rng = ActiveDocument.Bookmarks("CodeSnippet").Range
> rng.Copy
> MsgBox "Code copied to clipboard!"
> End Sub
8. Press **Ctrl + S** to save the Code.

#### Insert a Bookmark named CodeSnippet

This step tells Word which text (your code) should be copied when the button is clicked.

How to do it:

1. Select the **code** in your document that you want the button to **copy**.


2. Go to the **Insert tab** on the ribbon.


3. Click **Bookmark** (it's in the "Links" group).


4. In the Bookmark name field, type:

**CodeSnippet**


5. Click **Add**.
