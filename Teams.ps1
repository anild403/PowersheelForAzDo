# This script assumes that you move your mouse within the normal idletime (about 5 minutes).
Add-Type -AssemblyName System.Windows.Forms
$loopsToPreventIdle = 60 # works for this loops to prevent the idle. Means: After 60++ Minutes (e.g.) the computer goes idle. If you are using this script permanently, set it to 9999 or something like that
$secBevorIdle = 60*3 # if your idle time is 5 Minutes, set the variable to 60*5 to get the seconds of 5 minutes

$myshell = New-Object -com "Wscript.Shell" # starting  wscript shell to interact with an application
$app = start-process -filepath "notepad.exe" -passthru | Select-Object * #starting the application (notepad)

# you can also use the MainWindowTitle of the notepad (or whatever) instead of $myshell.AppActivate($app.Id) use $myshell.AppActivate($MainWindowTitle)
#$appProcess = get-process -id $app.Id #load the process again
#$MainWindowTitle = $appProcess.MainWindowTitle #$MainWindowTitle = "Unbenannt - Editor" # if you are using this, i recommend you check before sending a key if this is not empty
#write-host "MainWindowTitle: $MainWindowTitle"

start-sleep 2 # this is not neccessary, but usefull, to see that a notepad opens 
$myshell.AppActivate($app.Id) # get the application windows focused
$myshell.sendkeys("[$(get-date)] starting idle prevention. Keep this window open{ENTER}") # This is helpful so you dont get the idea of closing the window :)
$p1 = [System.Windows.Forms.Cursor]::Position # get the position of our mouse arrow
for ($i = 0; $i -lt $loopsToPreventIdle; $i++) {

    Start-Sleep -Seconds ($secBevorIdle-30) # wait a little bit. Normally you will move your mouse in this time... If not maybe you are watching a video or texting... or you are really idle
    $p2 = [System.Windows.Forms.Cursor]::Position # get the new position of your mouse to 

    if($p1.X -eq $p2.X -and $p1.Y -eq $p2.Y) {
        "[$(get-date)] Your mouse did not move (X: $($p1.X)) Y: $($p1.Y))"
        write-host "[$(get-date)] Sending key"
        $myshell.AppActivate($app.Id) # if the AppActivate not work (result = false), the script will send the keys (see the next two lines) to the currently focused window
        start-sleep 1
        $myshell.sendkeys("[$(get-date)]{ENTER}")
    } else {
        write-host "[$(get-date)] Your mouse moved"
    }
    Start-Sleep -Seconds 10  # wait a little bit
    
    $p1 = [System.Windows.Forms.Cursor]::Position # get the new position of your mouse
}
