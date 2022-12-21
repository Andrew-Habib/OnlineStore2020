
'Andrew Habib
'Date Due: Tuesday, June 16, 2020
'Course Code: ICS 2O1 - 02 (Computer Studies Grade 10 Course)
'Mrs. Tadros

'Program: Canada Fine Threads ©. Sample Online Store for High Quality Clothing (CPT)   
'Purpose: The purpose of the program is to create an easy - to - use, small yet conveniant programmed online store during the current Covid 19 Pandemic



Option Explicit On
Public Class frmOnlineStoreCPT

    'Declare the Global Variables (btnBuyShoes1 code)
    Dim intQuantity1 As Integer
    Dim intPrice1 As Integer

    'Declare the Global Variables (btnBuyShoes2 code)
    Dim intQuantity2 As Integer
    Dim intPrice2 As Integer

    'Declare the Global Variables (btnBuyShirt1 code)
    Dim intQuantity3 As Integer
    Dim intPrice3 As Integer

    'Declare the Global Variables (btnBuyShirt2 code)
    Dim intQuantity4 As Integer
    Dim intPrice4 As Integer

    'Declare the Global Variables (btnBuyPants1 code)
    Dim intQuantity5 As Integer
    Dim intPrice5 As Integer

    'Declare the Global Variables (btnBuyPants2 code)
    Dim intQuantity6 As Integer
    Dim intPrice6 As Integer

    'Declare the Global Variables (btnBuySweater1 code)
    Dim intQuantity7 As Integer
    Dim intPrice7 As Integer

    'Declare the Global Variables (btnBuySweater2 code)
    Dim intQuantity8 As Integer
    Dim intPrice8 As Integer

    'Declare the Global Variables (Total Cost code)
    Dim intTotalQ As Integer
    Dim intTotalP As Integer


    Private Sub frmOnlineStoreCPT_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Display a message box that greets the user once he clicks into the online store
        MessageBox.Show("Welcome to one of Canada's newest online clothing stores. 
Canada Fine Threads is at your service to provide you with the highest quality clothing at reasonable prices!", "Welcome")

        'Make sure timers are enabled
        tmrDate.Enabled = True
        tmrTime.Enabled = True


        'Initialize the counters
        intQuantity1 = 0
        intQuantity2 = 0
        intQuantity3 = 0
        intQuantity4 = 0
        intQuantity5 = 0
        intQuantity6 = 0
        intQuantity7 = 0
        intQuantity8 = 0
        intTotalQ = 0


        intPrice1 = 0
        intPrice2 = 0
        intPrice3 = 0
        intPrice4 = 0
        intPrice5 = 0
        intPrice6 = 0
        intPrice7 = 0
        intPrice8 = 0
        intTotalP = 0



    End Sub

    Private Sub tmrTime_Tick(sender As Object, e As EventArgs) Handles tmrTime.Tick

        'Display the current time while user is on the interface

        'Declare a local variable for the time (String)
        Dim strTime As String

        'Set time to current time
        strTime = Now
        Me.lblTime.Text = Format(Now, "hh:mm:ss")


    End Sub

    Private Sub tmrDate_Tick(sender As Object, e As EventArgs) Handles tmrDate.Tick

        'Display the current Date while the user is on the interface

        'Declare a local variable for the Date (String)
        Dim strDate As String

        'Set the Date to Today's Date
        strDate = Today
        Me.lblDate.Text = strDate

    End Sub



    Private Sub btnBuyShoes1_Click(sender As Object, e As EventArgs) Handles btnBuyShoes1.Click

        'Increment the counter for the Quantity using Global Variable listed at the top
        intQuantity1 = intQuantity1 + 1

        'Display the quantity based on number of clicks
        lblQuantity1.Text = intQuantity1

        'Increment the counter for the price using Global Variable listed at the top
        intPrice1 = intPrice1 + 20

        'Display the price based on number of clicks
        lblPrice1.Text = intPrice1

    End Sub

    Private Sub btnBuyShoes2_Click(sender As Object, e As EventArgs) Handles btnBuyShoes2.Click

        'Increment the counter for the Quantity using Global Variable listed at the top
        intQuantity2 = intQuantity2 + 1

        'Display the quantity based on number of clicks
        lblQuantity2.Text = intQuantity2

        'Increment the counter for the price using Global Variable listed at the top
        intPrice2 = intPrice2 + 25

        'Display the price based on number of clicks
        lblPrice2.Text = intPrice2

    End Sub

    Private Sub btnBuyShirt1_Click(sender As Object, e As EventArgs) Handles btnBuyShirt1.Click

        'Increment the counter for the Quantity using Global Variable listed at the top
        intQuantity3 = intQuantity3 + 1

        'Display the quantity based on number of clicks
        lblQuantity3.Text = intQuantity3

        'Increment the counter for the price using Global Variable listed at the top
        intPrice3 = intPrice3 + 15

        'Display the price based on number of clicks
        lblPrice3.Text = intPrice3

    End Sub

    Private Sub btnBuyShirt2_Click(sender As Object, e As EventArgs) Handles btnBuyShirt2.Click

        'Increment the counter for the Quantity using Global Variable listed at the top
        intQuantity4 = intQuantity4 + 1

        'Display the quantity based on number of clicks
        lblQuantity4.Text = intQuantity4

        'Increment the counter for the price using Global Variable listed at the top
        intPrice4 = intPrice4 + 15

        'Display the price based on number of clicks
        lblPrice4.Text = intPrice4

    End Sub

    Private Sub btnBuyPants1_Click(sender As Object, e As EventArgs) Handles btnBuyPants1.Click

        'Increment the counter for the Quantity using Global Variable listed at the top
        intQuantity5 = intQuantity5 + 1

        'Display the quantity based on number of clicks
        lblQuantity5.Text = intQuantity5

        'Increment the counter for the price using Global Variable listed at the top
        intPrice5 = intPrice5 + 20

        'Display the price based on number of clicks
        lblPrice5.Text = intPrice5

    End Sub

    Private Sub btnBuyPants2_Click(sender As Object, e As EventArgs) Handles btnBuyPants2.Click

        'Increment the counter for the Quantity using Global Variable listed at the top
        intQuantity6 = intQuantity6 + 1

        'Display the quantity based on number of clicks
        lblQuantity6.Text = intQuantity6

        'Increment the counter for the price using Global Variable listed at the top
        intPrice6 = intPrice6 + 30

        'Display the price based on number of clicks
        lblPrice6.Text = intPrice6

    End Sub

    Private Sub btnBuySweater1_Click(sender As Object, e As EventArgs) Handles btnBuySweater1.Click

        'Increment the counter for the Quantity using Global Variable listed at the top
        intQuantity7 = intQuantity7 + 1

        'Display the quantity based on number of clicks
        lblQuantity7.Text = intQuantity7

        'Increment the counter for the price using Global Variable listed at the top
        intPrice7 = intPrice7 + 25

        'Display the price based on number of clicks
        lblPrice7.Text = intPrice7

    End Sub

    Private Sub btnBuySweater2_Click(sender As Object, e As EventArgs) Handles btnBuySweater2.Click

        'Increment the counter for the Quantity using Global Variable listed at the top
        intQuantity8 = intQuantity8 + 1

        'Display the quantity based on number of clicks
        lblQuantity8.Text = intQuantity8

        'Increment the counter for the price using Global Variable listed at the top
        intPrice8 = intPrice8 + 25

        'Display the price based on number of clicks
        lblPrice8.Text = intPrice8

    End Sub

    Private Sub btnTotal_Click(sender As Object, e As EventArgs) Handles btnTotal.Click

        'Declare Local Constants
        Const dbl10 As Double = 0.9
        Const dbl25 As Double = 0.75
        Const dbl45 As Double = 0.55

        'Calculate the total cost and quantity using the Global Variables
        intTotalP = intPrice1 + intPrice2 + intPrice3 + intPrice4 + intPrice5 + intPrice6 + intPrice7 + intPrice8
        intTotalQ = intQuantity1 + intQuantity2 + intQuantity3 + intQuantity4 + intQuantity5 + intQuantity6 + intQuantity7 + intQuantity8

        'Use Select Case to check for discount on Total 
        Select Case intTotalQ

            Case = 1
                MessageBox.Show("Purchase more items to earn discounts!", "Disclaimer")

            Case 2 To 3
                MessageBox.Show("You get 10% off", "Message")
                intTotalP = intTotalP * dbl10

            Case 4 To 6
                MessageBox.Show("You get 25% off", "Message")
                intTotalP = intTotalP * dbl25

            Case >= 7
                MessageBox.Show("You get 45% off", "Message")
                intTotalP = intTotalP * dbl45

            Case Else
                MessageBox.Show("That is invalid input please try again", "Error")

        End Select

        'Display the Totals in the Labels
        Me.lblTotalP.Text = intTotalP
        Me.lblTotalQ.Text = intTotalQ

    End Sub
    Private Sub btnDiscount_Click(sender As Object, e As EventArgs) Handles btnDiscount.Click


        'Declare variables    
        Dim intCounterDiscount As Integer
        Const dblSaving As Double = 0.75
        Dim strTitle As String
        Dim strPrompt As String
        Dim strDiscountCode As String
        Dim dblDiscount As Double

        'Initialize the counter    
        intCounterDiscount = 0

        'Assign text to string variables    
        strTitle = "Enter Discount Code"
        strPrompt = "Please enter the Discount code below for 25 % OFF on all items."

        'Repeat the appearance of the input box 
        Do
            'Get the user input        
            strDiscountCode = InputBox(strPrompt, strTitle)

            'Increment the Counter  
            intCounterDiscount = intCounterDiscount + 1

            dblDiscount = Val(lblTotalP.Text) * dblSaving

            'Check if the password is correct  
        Loop While strDiscountCode <> "SAVE25" And intCounterDiscount < 5

        If strDiscountCode <> "SAVE25" And intCounterDiscount = 5 Then
            MessageBox.Show("Too Many Failed Attempts!", “Error”)

        Else

            'If the password is correct display a message, fulfill the discount and remove the Discount Button
            MessageBox.Show("You get 25% OFF extra on all items!", "Discount")
            Me.btnDiscount.Visible = False
            Me.lblNewTotalP.Text = dblDiscount

        End If

    End Sub
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click

        'Reset the contents in the checkout area
        Me.lblQuantity1.Text = Nothing
        Me.lblQuantity2.Text = Nothing
        Me.lblQuantity3.Text = Nothing
        Me.lblQuantity4.Text = Nothing
        Me.lblQuantity5.Text = Nothing
        Me.lblQuantity6.Text = Nothing
        Me.lblQuantity7.Text = Nothing
        Me.lblQuantity8.Text = Nothing
        Me.lblTotalQ.Text = Nothing

        Me.lblPrice1.Text = Nothing
        Me.lblPrice2.Text = Nothing
        Me.lblPrice3.Text = Nothing
        Me.lblPrice4.Text = Nothing
        Me.lblPrice5.Text = Nothing
        Me.lblPrice6.Text = Nothing
        Me.lblPrice7.Text = Nothing
        Me.lblPrice8.Text = Nothing
        Me.lblTotalP.Text = Nothing
        Me.lblNewTotalP.Text = Nothing

        intQuantity1 = 0
        intQuantity2 = 0
        intQuantity3 = 0
        intQuantity4 = 0
        intQuantity5 = 0
        intQuantity6 = 0
        intQuantity7 = 0
        intQuantity8 = 0
        intTotalQ = 0


        intPrice1 = 0
        intPrice2 = 0
        intPrice3 = 0
        intPrice4 = 0
        intPrice5 = 0
        intPrice6 = 0
        intPrice7 = 0
        intPrice8 = 0
        intTotalP = 0


        'Initialize the counter
    End Sub
    Private Sub btnPay_Click(sender As Object, e As EventArgs) Handles btnPay.Click

        'Display a message Box
        MessageBox.Show("Paid. Hope you enjoyed your service. You will recieve your items in up to 30 days.", "Thank you")

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click

        'Message before the user terminates the program
        MessageBox.Show("Thank you for shopping at Canada Fine Threads!")

        'Terminate the Application
        Application.Exit()

    End Sub


End Class
