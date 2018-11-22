# Console and logging
## Invoke-ColorOutput
* Author: Marc Dekeyser
* DESCR: Provides a way to output colorized text to the console without using write-host. Thus avoiding the killing of puppies. Function will also reapply the original color settings once completed.
* PARAM: OBJECT (Alias Message, MSG)
* PARAM: ForeGroundColor (Alias Fore)
* PARAM: BackGroundColor (Alias back, bgr)
* PARAM: NoNewLine

## Invoke-ProgressHelper
* Author: Unknown. Modified this to suit my needs.
* PARAM: StepNumber
* PARAM: Status message
* PARAM: ID (place on the screen. )
* PARAM: Activity
* PARAM: Steps (Total number of steps)

## Invoke-WriteLog
* Author: Marc Dekeyser
* DESCR: Writes timestamped, colorized output to the screen and adds the same to a logfile if the "Runlog" parameter is specified. Color is determined by the prefixes "SUCCESS:", "STATUS:", "INFO:", "ALERT:", "ERROR:", "AUDIT:", "NEWLINE", "TITLE:", "TEXT:", "PROMPT:". Note that this function does not use invoke-ColorOutput, but used write-host, thus it kills puppies and does not output object based formatting. This script will likely to be rewritten.
* PARAM: Message (Required)
* PARAM: Runlog (Optional)

## New-BalloonTip
* Author: Marc Dekeyser, based on the multitude of balloon tip scripts out there...
* DESCR: Creates a balloontip notification...
* PARAM: Title
* PARAM: Message
* PARAM: ToolTipIcon
* PARAM: Duration

## New-PopUp
* Author: Marc Dekeyser
* DESCR: Creates a new pop up WScript.Shell
* PARAM: Message
* PARAM: Title
* PARAM: Time (Time to display)
* PARAM: Buttons
* PARAM: Icon
* PARAM: DefaultButton
