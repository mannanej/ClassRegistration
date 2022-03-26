##################################################################################################################################################################
#    This script will register you for classes at 7:30 AM
#    To run this script, you must first install Selenium and Chrome Driver for PowerShell
#    (In case you couldn't tell from 'Chrome Driver', this will only run with Google Chrome, so you will need that too)
#    Install Selenium ----- https://github.com/adamdriscoll/selenium-powershell ----- Run the command in Command Prompt
#    Install ChromeDriver ----- https://chromedriver.storage.googleapis.com/index.html?path=75.0.3770.140/ -----
#    Make User Account Admin ----- Add-LocalGroupMember -Group “Administrators” -Member “YOUR_LAPTOP_USERNAME_HERE” ----- Run this command in PowerShell (Admin)
#    (Once you use this command to make yourself Admin, your PC will need to restart for it to take affect)
#                  ----------Copyright 2019, Eddie Mannan, All rights reserved----------
#    (Ok, I didn't really copyright any of this, but come on; if you are using this, you can at least give me the credit. This took more effort to make than you would think. Thanks :))
##################################################################################################################################################################

##################################################################################################################################################################
# Im going to put all of my variables in this section, so they are all loaded in before the time, saving time. Get it? Good.
# Change the username to your BannerWeb username, and the same with the password. The PinNumber is the 6-digit pin from your advisor. The class codes are the CRNs for the classes you want.
# If you are taking more than 4 classes, remove the '#' in front of the green text, and add the CRN. ALSO, uncomment the corresponding lines in function Class-Registration.

$Username = "USERNAME"
$Password = "PASSWORD"

$PinNumber = 123456

$ClassCode1 = 1234
$ClassCode2 = 1234
$ClassCode3 = 1234
$ClassCode4 = 1234
#$ClassCode5 = 1234
#$ClassCode6 = 1234
#$ClassCode7 = 1234
#$ClassCode8 = 1234

$TargetTime = Get-Date -Hour 7 -Minute 30 -Second 00 -Millisecond 100
##################################################################################################################################################################

Function Class-Registration {

# This function will type in your desired CRNs, and hit 'Submit' once they are all entered. Here is where you would uncomment for more than 4 classes.
# This first submit is the one that hits the button once its 7:30.

    # Click Submit
    Find-SeElement -Driver $Driver -XPath "//input[@type='submit']" | Invoke-SeClick -Driver $Driver -JavaScriptClick

    # Register for the classes
    $Class1 = Find-SeElement -Driver $Driver -Id "crn_id1"
    Send-SeKeys -Element $Class1 -Keys "$ClassCode1"

    $Class2 = Find-SeElement -Driver $Driver -Id "crn_id2"
    Send-SeKeys -Element $Class2 -Keys "$ClassCode2"

    $Class3 = Find-SeElement -Driver $Driver -Id "crn_id3"
    Send-SeKeys -Element $Class3 -Keys "$ClassCode3"

    $Class4 = Find-SeElement -Driver $Driver -Id "crn_id4"
    Send-SeKeys -Element $Class4 -Keys "$ClassCode4"

    #$Class5 = Find-SeElement -Driver $Driver -Id "crn_id5"
    #Send-SeKeys -Element $Class5 -Keys "$ClassCode5"

    #$Class6 = Find-SeElement -Driver $Driver -Id "crn_id6"
    #Send-SeKeys -Element $Class6 -Keys "$ClassCode6"

    #$Class7 = Find-SeElement -Driver $Driver -Id "crn_id7"
    #Send-SeKeys -Element $Class7 -Keys "$ClassCode7"

    #$Class8 = Find-SeElement -Driver $Driver -Id "crn_id8"
    #Send-SeKeys -Element $Class8 -Keys "$ClassCode8"

    # Click Submit
    Send-SeKeys -Element $Class1 -Keys ([OpenQA.Selenium.Keys]::Enter)

    # Lets go make make a check-up text box
    Check-Everything
}

Function Compare-Time {

# This function checks what time it currently is once every millisecond, then once the targeted time above is reached (7:30:00:100), it will go to Class-Registration.

    $CurrentTime = Get-Date

    While ($CurrentTime -notlike $TargetTime) {

    $CurrentTime = Get-Date

    }

    Class-Registration
}


Function GoTo-Website {
    
# This function navagates you to Banner Web, takes you to the Registration page, and enters your pin. Then, it goes to Compare-Time, and waits untill 7:30.

    $Driver = Start-SeChrome
    $Driver.Manage().Window.Size = [System.Drawing.Size]::new(2000, 800)
    $Driver.Manage().Timeouts().ImplicitWait = [TimeSpan]::FromSeconds(10)
    Enter-SeUrl -Driver $Driver -Url "https://bxess-prod-hv.rose-hulman.edu/BanSS/twbkwbis.P_GenMenu?name=bmenu.P_StuMainMnu"

    # Log-In to banner web
    $UN = Find-SeElement -Driver $Driver -Name 'sid'
    $PS = Find-SeElement -Driver $Driver -Name 'PIN'
    Send-SeKeys -Element $UN -Keys "$Username"
    Send-SeKeys -Element $PS -Keys "$Password"
    Send-SeKeys -Element $PS -Keys ([OpenQA.Selenium.Keys]::Enter)

    # Select the Student tab
    Find-SeElement -Driver $Driver -LinkText 'Student' | Invoke-SeClick -Driver $Driver

    # Select the registration option
    Find-SeElement -Driver $Driver -LinkText 'Registration, Individual Class Schedule, Schedule of Classes By Term (Live Data)' | Invoke-SeClick -Driver $Driver

    # Select the Register for Class option
    Find-SeElement -Driver $Driver -LinkText 'Register for Classes' | Invoke-SeClick -Driver $Driver

    # Submit the quarter selection
    Find-SeElement -Driver $Driver -XPath "//input[@type='submit']" | Invoke-SeClick -Driver $Driver -JavaScriptClick

    # Plug in the PIN
    $PIN = Find-SeElement -Driver $Driver -Name 'pin'
    Send-SeKeys -Element $PIN -Keys "$PinNumber"

    # Now, wait till its time
    Compare-Time

}

Function Check-Everything {

    # This will set up the basic text box and throw it on the screen

    $TextBox = [System.Windows.MessageBox]::Show('Did you get all of your classes? (I hope so :p)',' Class Registration Script Successfull','YesNo','Question')

    # This will tell the buttons what to do
    switch ($TextBox) {

        # This tells the "Yes" button to close the PowerShell and CMD windows and the Chrome browser
        'Yes' {

            Stop-SeDriver -Driver $Driver
            Stop-Process -Id $PID
            
        }

        # This tells the "No" button to stop running the PowerShell script but leaves Chrome open for manually entering CRN's
        'No' {

            Exit

        }
    }
}

# This is the first "Real Line" of code. Once you hit RUN, this will get the ball rolling

GoTo-Website
