#####################################################################
### File created by Thomas Gruber, 2018                           ###
#####################################################################

from distutils.util import strtobool
from robot.api import logger
import time

class RanorexLibrary(object):
    """ The RanorexLibrary main object.

    It is imported into a Robot test suite file in the "Library" section:

    *** Settings***
    Documentation           This is the file where the RanorexLibrary will be imported.
    Library                 RanorexLibrary      path\\to\\Ranorex

    The RanorexLibrary takes one argument:
    The path to Ranorex has to be given so the RobotLibrary knows where to import the Ranorex .dll files from. Normally this path looks something like this: C:\\Program Files (x86)\\Ranorex 8.3Beta. Please make sure to use double back slashes (because of Robot-reasons).
    """

    __version__ = '0.1'

    ROBOT_LIBRARY_SCOPE = 'TEST_SUITE'
    ROBOT_LIBRARY_DOC_FORMAT = 'reST'

    _logLevel = "INFO"

    def __init__(self, pathToRanorex):
        import setupRanorexLibrary
        setupRanorexLibrary.importDlls(pathToRanorex)

        global Ranorex
        import Ranorex
        global System
        import System

        Ranorex.Core.Resolver.AssemblyLoader.Initialize()
        Ranorex.TestingBootstrapper.SetupCore()

        Ranorex.Mouse.DefaultMoveTime = 300
        Ranorex.Keyboard.DefaultKeyPressTime = 100
        Ranorex.Delay.SpeedFactor = 1

    def _log(self, msg):
        logger.write(msg, self._logLevel, html=False)

    def run_application(self, appname, arguments = "", workingDirectory = "", maximized = "False"):
        """ Runs an Application.

        This is the suggested way to run an application using Ranorex functionality. Other libraries might offer other functions that open an application, so you have to decide which one you like the most.

        :param appname: This is the path to the executable file of the application to be started.
        :param arguments: This argument is passed to the started application as command line arguments.
        :param workingDirectory: This is the path to the directory that Ranorex tries to give the application as working directory.
        :param maximized: True or False. Whether Ranorex tries to open the application with a maximized window or not. Might not work for all applications.

        Example:
        | `Run Application` | calc.exe |
        | `Run Application` | C:\\Program Files\\Internet Explorer\\iexplore.exe |  |  | True |
        | `Run Application` | yourApp.exe | /help | C:\\path\\to\\yourWorkingDirectory | False |
        """
        self._log("Starting application " + appname + ".")
        maxim = False
        if maximized == "True":
            maxim = True
        Ranorex.Host.Local.RunApplication(appname, arguments, workingDirectory, maxim)

    def close_application(self, ranorexpath, gracePeriod = "0"):
        """ Closes an application that contains a specified UI element.

        This keyword looks for a UI element specified by a RanoreXPath and tries to close the parent process of this element.

        :param ranorexpath: This path specifies an element within the application that should be closed.
        :param gracePeriod: Milliseconds until the application is killed if it hasn't closed properly until then. If this value is 0, the app will never be killed.

        :returns: True if the application has closed within the grace period, otherwise false.

        Example:
        | `Close Application` | /winapp[@packagename='Microsoft.WindowsCalculator'] |  |
        | `Close Application` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] | 300 |
        """
        self._log("Closing application with element " + ranorexpath + " within " + gracePeriod + "ms.")
        intGracePeriod = int(gracePeriod)
        return Ranorex.Host.Current.CloseApplication(ranorexpath, intGracePeriod)

    def click(self, ranorexpath, location = "Center", mousebutton = "Left", duration = "Ranorex.Mouse.DefaultMoveTime", count = "1"):
        """ Performs a mouse click on a UI element.

        :param ranorexpath: This is the RanoreXPath of the element that gets clicked.
        :param location: The location where the element should be clicked. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight
        :param mousebutton: Which mouse button should be clicked. Possible values are listed on the corresponding .NET framework page: https://msdn.microsoft.com/de-de/library/system.windows.forms.mousebuttons(v=vs.110).aspx
        :param duration: The duration of the mouse click in ms. Defaults value is 300ms.
        :param count: Number of clicks that should be performed. Default is (of course) 1.

        Example:
        | `Click` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] |  |  |  |  |
        | `Click` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] | UpperLeft |  |  |  |
        | `Click` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] | UpperLeft |  | 350 |  |
        | `Click` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] |  |  |  | 2 |
        """
        self._log("Clicking on element " + ranorexpath + " at location " + location + " " + count + " time(s) with " + mousebutton + " mouse button, taking " + duration + " ms.")
        if mousebutton == "":
            mousebutton = "Left"
        if location == "":
            location = "Center"
        if count == "":
            count = "1"
        if duration == "":
            duration = "Ranorex.Mouse.DefaultMoveTime"
        self._click(ranorexpath, location, mousebutton, duration, count)

    def _normalizeMouseButton(self, mousebutton):
        return "System.Windows.Forms.MouseButtons." + mousebutton

    def _normalizeLocation(self, location):
        return "Ranorex.Location." + location

    def _click(self, ranorexpath, location, mousebutton, duration, count):
        mousebutton = self._normalizeMouseButton(mousebutton)
        location = self._normalizeLocation(location)
        exec("Ranorex.Unknown(ranorexpath).Click(" + mousebutton + ", " + location + ", " + "int(" + count + "), " + duration + ")")

    def right_click(self, ranorexpath, location = "Center", duration = "Ranorex.Mouse.DefaultMoveTime", count = "1"):
        """ Performs a right click on a UI element.

        This action is equivalent to the click action if "Right" is given as the mousebutton parameter.

        :param ranorexpath: This is the RanoreXPath of the element that gets clicked.
        :param location: The location where the element should be clicked. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight
        :param duration: The duration of the mouse click in ms. Defaults value is 300ms.
        :param count: Number of clicks that should be performed. Default is (of course) 1.

        Example:
        | `Right Click` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] |  |  |  |
        | `Right Click` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] | UpperLeft |  |  |
        | `Right Click` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] | UpperLeft | 350 |  |
        | `Right Click` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] |  |  | 2 |
        """
        self._log("Right clicking on element " + ranorexpath + " at location " + location + " " + count + " time(s) , taking " + duration + " ms.")
        if location == "":
            location = "Center"
        if count == "":
            count = "1"
        if duration == "":
            duration = "Ranorex.Mouse.DefaultMoveTime"
        self._click(ranorexpath, location, "Right", duration, count)

    def double_click(self, ranorexpath, location = "Center", mousebuttons = "Left", duration = "Ranorex.Mouse.DefaultMoveTime"):
        """ Performs a double click on a UI element.

        :param ranorexpath: This is the RanoreXPath of the element that gets clicked.
        :param location: The location where the element should be clicked. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight
        :param mousebutton: Which mouse button should be clicked. Possible values are listed on the corresponding .NET framework page: https://msdn.microsoft.com/de-de/library/system.windows.forms.mousebuttons(v=vs.110).aspx
        :param duration: The duration of the mouse click in ms. Defaults value is 300ms.

        Example:
        | `Double Click` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] |  |  |  |
        | `Double Click` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] | UpperLeft | Right | 100 |
        """
        self._log("Double clicking on element " + ranorexpath + " at location " + location + " with " + mousebuttons + " mouse button, taking " + duration + " ms.")
        if location == "":
            location = "Center"
        if mousebuttons == "":
            mousebuttons = "Left"
        if duration == "":
            duration = "Ranorex.Mouse.DefaultMoveTime"
        mousebuttons = self._normalizeMouseButton(mousebuttons)
        location = self._normalizeLocation(location)
        exec("Ranorex.Unknown(ranorexpath).DoubleClick(" + mousebuttons + ", " + location + ", " + duration + ")")

    def _moveMouseToElement(self, ranorexpath, location, duration):
        exec("Ranorex.Unknown(ranorexpath).MoveTo(" + location + ", " + duration + ")")

    def mouse_down(self, ranorexpath, location = "Center", button = "Left", duration = "Ranorex.Mouse.DefaultMoveTime"):
        """ Performs a Mouse Down action on a UI element.

        :param ranorexpath: This is the RanoreXPath of the element that gets clicked.
        :param location: The location where the element should be clicked. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight
        :param mousebutton: Which mouse button should be clicked. Possible values are listed on the corresponding .NET framework page: https://msdn.microsoft.com/de-de/library/system.windows.forms.mousebuttons(v=vs.110).aspx
        :param duration: The duration of the mouse click in ms. Defaults value is 300ms.

        Example:
        | `Mouse Down` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] |  |  |  |
        | `Mouse Down` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] | LowerCenter | Right | 250 |
        """
        self._log("Mouse down on element " + ranorexpath + " at location " + location + " with " + button + " mouse button, taking " + duration + " ms.")
        if location == "":
            location = "Center"
        if button == "":
            button = "Left"
        if duration == "":
            duration = "Ranorex.Mouse.DefaultMoveTime"
        location = self._normalizeLocation(location)
        button = self._normalizeMouseButton(button)
        self._moveMouseToElement(ranorexpath, location, duration)
        exec("Ranorex.Mouse.ButtonDown(" + button + ")")

    def mouse_move(self, ranorexpath, location = "Center", duration = "Ranorex.Mouse.DefaultMoveTime"):
        """ Moves the mouse cursor the the specified location.

        :param ranorexpath: This is the RanoreXPath of the element the mouse cursor is moved to.
        :param location: The location where the element should be clicked. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight
        :param duration: The duration of the mouse click in ms. Defaults value is 300ms.

        Example:
        | `Mouse Move` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] |  |  |
        | `Mouse Move` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num8Button'] | UpperRight | 150 |
        """
        self._log("Move mouse cursor to element " + ranorexpath + " at location " + location + ", taking " + duration + " ms.")
        if location == "":
            location = "Center"
        if duration == "":
            duration = "Ranorex.Mouse.DefaultMoveTime"
        location = self._normalizeLocation(location)
        self._moveMouseToElement(ranorexpath, location, duration)

    def mouse_up(self, ranorexpath = "", button = "Left", location = "Center", duration = "Ranorex.Mouse.DefaultMoveTime"):
        """ Performs a mouse up action. Moves the cursor the desired location if given.

        :param ranorexpath: This is the RanoreXPath of the element that gets clicked. If no path is given, then the mouse up is performed on the current cursor location.
        :param mousebutton: Which mouse button should be clicked. Possible values are listed on the corresponding .NET framework page: https://msdn.microsoft.com/de-de/library/system.windows.forms.mousebuttons(v=vs.110).aspx
        :param location: The location where the element should be clicked. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight
        :param duration: The duration of the mouse click in ms. Defaults value is 300ms.

        Example:
        | `Mouse Up` |  |  |  |  |
        | `Mouse Up` | [@packagename='Microsoft.WindowsCalculator']//button[@automationid='num2Button'] |  |  |  |
        | `Mouse Up` | [@packagename='Microsoft.WindowsCalculator']//button[@automationid='num2Button'] | Right | UpperLeft | 100 |
        """
        self._log("Mouse up on element " + ranorexpath + " at location " + location + " with " + button + " mouse button, taking " + duration + " ms.")
        if button == "":
            button = "Left"
        button = self._normalizeMouseButton(button)
        if location == "":
            location = "Center"
        location = self._normalizeLocation(location)
        if duration == "":
            duration = "Ranorex.Mouse.DefaultMoveTime"
        if ranorexpath != "":
            self._moveMouseToElement(ranorexpath, location, duration)
        exec("Ranorex.Mouse.ButtonUp(" + button + ")")

    def key_shortcut(self, sequence):
        """ Performs a key shortcut specified by a string representation.

        Take care that the GUI is in the desired state already, the Key Shortcut keyword does not for anything to be loaded fully. Thus, an explicit waiting keyword might be necessary before performing a key shortcut.

        Each key is represented by a single character or an escape group. To specify a single keyboard character, use the character itself (e.g. use "aBc" to press the keys A, B+Shift, and C after another). Only the '{' character has a special meaning and needs to be escaped by preceding it with another '{' (specify "{{" to issue a '{' key press).
        Escape groups, signaled by braces "{}", may be used to produce a key action with any of the keys specified by the Keys enumeration (https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.keys?redirectedfrom=MSDN&view=netframework-4.7.2). "{Z}" means that the 'z' key should be pressed, and "{return}" or "{enter}" that the Return key is to be pressed.

        Additionally, in an escape group you can specify a modifier that determines whether the key should be hold down (e.g. {CONTROL down}), released (e.g. {shift up}), or pressed a number of times (e.g. {z 3}). In an escape group, key name and modifier must be separated by a single shift character.

        The keys A to Z set the key modifiers (ALT, CTRL, SHIFT) and in particular the shift key depending on their case, even if used in an escape group. If these keys should not modify the shift key state, wrap them into an escape group and add "key" to the character, e.g. "{Rkey}" to press the R key without changing the shift key state.

        :param sequence: Sequence that represents the key sequence to press.

        Example:
        | `Key Shortcut` | {F12} | # Presses the F12 key
        | `Key Shortcut` | {RMenu down}{qKey}{RMenu up} | # Types the @ symbol on german keyboards
        | `Key Shortcut` | {Control down}{cKey}{Control up} | # Copying from C/P
        """
        self._log("Type key shortcut \"" + sequence + "\".")
        Ranorex.Keyboard.Press(sequence)

    def start_browser(self, url, browser, browserArgs = "", killExisting = "True", maximized = "False", clearCache = "False", incognitoMode = "False", clearCookies = "False"):
        """ Starts a Browser Instance.

        :param url: The URL to open on startup.
        :param browser: The browser to start. Supported browser names: IE, Edge, Firefox, Chrome, Chromium
        :param browserArgs: The command line arguments that are passed to the browser process.
        :param killExisting: Boolean value. Specifies whether an existing browser window should be closed before starting.
        :param maximized: Boolean value. Specifies whether the browser window should be startd in maximized mode.
        :param clearCache: Boolean value. Specifies whether the cache should be cleared before startup to ensure the same user experience for every test case.
        :param incognitoMode: Boolean value. Specifies whether the browser should be started in incognito mode. This normally means that Ranorex has no access to the dom tree anymore.
        :param clearCookies: Boolean value. Specifies whether the coockies should be cleared before startup.

        Examples:
        | `Start Browser` | www.ranorex.com | Firefox |  |  |  |  |  |
        | `Start Browser` | www.ranorex.com | Chrome |  | False | True | false | yes | No |
        """
        self._log("Start Browser " + browser + " at " + url + ".")
        if killExisting == "":
            killExisting = "True"
        if maximized == "":
            maximized = "False"
        if clearCache == "":
            clearCache = "False"
        if incognitoMode == "":
            incognitoMode = "False"
        if clearCookies == "":
            clearCookies = "False"
        Ranorex.Host.Current.OpenBrowser(url, browser, browserArgs, strtobool(killExisting), strtobool(maximized), strtobool(clearCache), strtobool(incognitoMode), strtobool(clearCookies))

    def close_browser(self, ranorexpath, gracePeriod = "0"):
        """ Closes a browser window.

        Internally uses the close application keyword. See there for full documentation.
        """
        self._log("Close browser with element " + ranorexpath + " within " + gracePeriod + " ms.")
        intGracePeriod = int(gracePeriod)
        return Ranorex.Host.Current.CloseApplication(ranorexpath, intGracePeriod)

    def wait_for(self, ranorexpath, duration = 30000):
        """ Waits for an element to exist.

        Since Ranorex doesn't know how to wait for something implicitely in many cases, this has to be done explicitely. A very typical example would be to wait for a UI element to come into existance (often due to the UI being fully loaded).

        :param ranorexpath: The RanoreXPath of the element that Ranorex waits for.
        :param duration: The duration in ms that Ranorex waits for the element. If it isn't found within the specified timeout, an error is raised.

        Example:
        | `Wait For` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num5Button'] |  |
        | `Wait For` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num5Button'] | 5000 |
        """
        self._log("Wait " + duration + "ms for element " + ranorexpath + " to be found.")
        intRanorexpath = Ranorex.Core.RxPath(ranorexpath)
        intDuration = Ranorex.Duration(int(duration))
        newElement = None

        elementFound, newElement = Ranorex.Host.Local.TryFindSingle(intRanorexpath, intDuration) # third parameter (an out parameter) is implicitely returned back as a second return value, thanks to ironpython. yay
        if not elementFound:
            raise AssertionError('Element hasn\'t been found within the specified timeout of ' + duration + 'ms: ' + ranorexpath)

    def get_attribute_value(self, ranorexpath, attribute):
        """ Returns an attribute value of a UI element as string.

        :param ranorexpath: RanoreXPath of the element that the attribute is read from.
        :param attribute: The attribute value that should be read.

        :returns: The value of the attribute as string value.

        Example:
        | ${retValue} | `Get Attribute Value` | /winapp[@packagename='Microsoft.WindowsCalculator']/?/?/text[@automationid='CalculatorResults']/container[@automationid='textContainer'] | Caption |
        """
        self._log("Get the value of the attribute " + attribute + " from element " + ranorexpath + ".")
        return Ranorex.Unknown(ranorexpath).GetAttributeValue[str](attribute)

    def set_attribute_value(self, ranorexpath, attribute, value):
        """ Sets an attribute value of a UI element.

        :param ranorexpath: RanoreXPath of the element that has the attribute changed.
        :param attribute: The attribute of the element that gets changed.
        :param value: The new value of the attribute.

        Example:
        | `Set Attribute Value` | /form[@controlname='RxMainFrame']//text[@accessiblename='Enter your name'] | AccessibleValue | Dr. Strange |
        """
        self._log("Set the attribute " + attribute + " of element " + ranorexpath + " to the value \"" + value + "\".")
        Ranorex.Unknown(ranorexpath).Element.SetAttributeValue(attribute, value)

    def key_sequence(self, ranorexpath, value):
        """ Enters a key sequence into a specified UI element.

        If you want to press keys without the necessity of a UI element, use the Key Shortcut keyword instead.

        Each key is represented by a single character or an escape group. To specify a single keyboard character, use the character itself (e.g. use "aBc" to press the keys A, B+Shift, and C after another). Only the '{' character has a special meaning and needs to be escaped by preceding it with another '{' (specify "{{" to issue a '{' key press).
        Escape groups, signaled by braces "{}", may be used to produce a key action with any of the keys specified by the Keys enumeration (https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.keys?redirectedfrom=MSDN&view=netframework-4.7.2). "{Z}" means that the 'z' key should be pressed, and "{return}" or "{enter}" that the Return key is to be pressed.

        Additionally, in an escape group you can specify a modifier that determines whether the key should be hold down (e.g. {CONTROL down}), released (e.g. {shift up}), or pressed a number of times (e.g. {z 3}). In an escape group, key name and modifier must be separated by a single shift character.

        The keys A to Z set the key modifiers (ALT, CTRL, SHIFT) and in particular the shift key depending on their case, even if used in an escape group. If these keys should not modify the shift key state, wrap them into an escape group and add "key" to the character, e.g. "{Rkey}" to press the R key without changing the shift key state.

        :param ranorexpath: RanoreXPath of the UI element that will receive the key sequence.
        :param sequence: Sequence that represents the key sequence to press.

        | `Key Sequence` | /form[@processname='iexplore' and @visible='True']/element[@accessiblename='Navigation Bar']//text[class='Edit'][1]  | www.ranorex.com |
        | `Key Sequence` | /form[@title='Untitled - Notepad']/text[@controlid='15'] | II. Do not fear difficulty. Hard ground makes stronger roots. |
        """
        self._log("Type key sequence \"" + value + "\" into element " + ranorexpath)
        Ranorex.Unknown(ranorexpath).PressKeys(value)

    def validate_attribute_equal(self, ranorexpath, attribute, value):
        """Validates that an attribute is equal to the specified value.

        Since a Ranorex Validation action requires a repository item to work on, this keyword is a simple version implemented directly in python. Therefore its functionality might not be on par with the Ranorex validations.

        :param ranorexpath: The RanoreXPath of the element that the validation works on.
        :param attribute: The attribute that should be validated.
        :param value: The value that the attribute should be validated against.

        :raises: AsstionError if the validation fails.

        Example:
        | `Validate Attribute Equal` | /form[@controlname='RxMainFrame']/?/?/tabpage[@controlname='RxTabIntroduction']/text[@controlname='lblWelcomeMessage'] | ControlText | Welcome, Dr. Strange! |
        """
        #Problem is: the ranorex validation action needs a repo item to work on, so I have to do it manually.
        varToVal = Ranorex.Unknown(ranorexpath).GetAttributeValue[str](attribute)
        if not varToVal == value:
            raise AssertionError("Elements are not equal. Expected " + value + ", but got " + varToVal + " instead.")

    def validate_attribute_not_equal(self, ranorexpath, attribute, value):
        """ Validates that an attribute is not equal to the specified value.

        See in the validate_attribute_equal keyword for a documentation is these two keywords are the same apart from a small "not".
        """
        varToVal = Ranorex.Unknown(ranorexpath).GetAttributeValue[str](attribute)
        if varToVal == value:
            raise AssertionError("Elements are equal, although they shouldn't be. Expected and actual value: " + value)

    def run_mobile_app(self, endpoint, appname, resetState = "True"):
        """ Starts an application on a mobile device.

        :param endpoint: Name of the endpoint that the application should be started on. This endpoint has to be added using the Add Device keyword first, usually.
        :param appname: This is the name of the app that has to be started. On Android, this name might look like this: com.dropbox.android
        :param resetState: True if the app should be restarted, False if it should just be brought into the foreground if it is running already.

        Example:
        | `Run Mobile App` | Nexus 9 | com.dropbox.android | False |
        """
        self._log("Run mobile application " + appname + " on endpoint " + endpoint + ".")
        if resetState == "":
            resetState = "True"
        Ranorex.Host.Local.RunMobileApp(endpoint, appname, strtobool(resetState))

    def add_device(self, name, platform, typeName, address): #this cost me a lot of sweat and tears
        """ Adds a device endpoint for testing iOS and Android applications.

        This keywords adds a device as an endpoint for test execution of mobile applications.

        :param name: This is the name of the endpoint that has to be used in the Run Mobile App keyword afterwards.
        :param platform: Has to be either Android or iOS depending on the device endpoint.
        :param typeName: Has to be either WLAN or USB, depending on how the device is connected to the machine that runs Robot.
        :param address: Has to be the IP address (if connected via WLAN) or the USB serial (if connected via USB). Both can be found easiest with using Ranorex Studio or Spy.

        Example:
        | `Add Device` | Galaxy S7 Test Device | Android | WLAN | 192.168.14.3 |
        | `Add Device` | iPad 10 Test Device | iOS | USB | HT4AWJT01500 |
        """
        self._log("Add " + platform + " device " + name + " via " + typeName + " with address " + address + ".")
        platform = "Ranorex.Core.Remoting.RemotePlatform." + platform
        typeName = "Ranorex.Core.Remoting.RemoteConnectionType." + typeName
        exec("Ranorex.Core.Remoting.RemoteServiceLocator.Service.AddDevice(\"" + name + "\", " + platform + ", " + typeName + ", \"" + address + "\")")

    def close_mobile_app(self, ranorexpath, gracePeriod = "0"):
        """ Closes an application that contains a specified UI element.

        This keyword looks for a UI element specified by a RanoreXPath and tries to close the parent process of this element.

        :param ranorexpath: This path specifies an element within the application that should be closed.
        :param gracePeriod: Milliseconds until the application is killed if it hasn't closed properly until then. If this value is 0, the app will never be killed.

        :returns: True if the application has closed within the grace period, otherwise false.

        Example:
        | `Close Application` | /winapp[@packagename='Microsoft.WindowsCalculator'] |  |
        | `Close Application` | /winapp[@packagename='Microsoft.WindowsCalculator']//button[@automationid='num1Button'] | 300 |
        """
        self._log("Closing application with element " + ranorexpath + " within " + gracePeriod + "ms.")
        intGracePeriod = int(gracePeriod)
        Ranorex.Host.Current.CloseApplication(ranorexpath, intGracePeriod)
    
    def touch(self, ranorexpath, location = "Center", duration = "Ranorex.Mouse.DefaultMoveTime"):
        """ This keyword performs a touch on a mobile element.

        In its core, this function is very similar to how clicks work in Ranorex.

        :param ranorexpath: RanoreXPath of the element that should get touched.
        :param location: The location where the element should be touched. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight
        :param duration: Duration of the touch in ms.

        Example:
        | `Touch` | /mobileapp[@title='com.dropbox.android']//button[@accessiblename='Enter'] |  |  |
        | `Touch` | /mobileapp[@title='com.dropbox.android']//button[@innertext='Cancel'] | UpperRight |  |
        | `Touch` | /mobileapp[@title='com.dropbox.android']//button[@accessiblename='Enter'] | CenterLeft | 1000 |
        """
        self._log("Touch element " + ranorexpath + " at location " + location + " for " + duration + "ms.")
        if location == "":
            location = "Center"
        if duration == "":
            duration = "Ranorex.Mouse.DefaultMoveTime"
        location = self._normalizeLocation(location)
        exec("Ranorex.Unknown(ranorexpath).Touch(" + location + ", " + duration + ")")

    def double_tap(self, ranorexpath, location = "Center"):
        """ This keyword performs a double tap on a mobile element.

        It is very similar to the double click action.

        :param ranorexpath: RanoreXPath of the element that should get double tapped.
        :param location: The location where the element should be touched. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight

        Example:
        | `Double Tap` | /mobileapp[@title='com.dropbox.android']//button[@accessiblename='Enter'] |  |
        | `Double Tap` | /mobileapp[@title='com.dropbox.android']//button[@accessiblename='Enter'] | CenterRight |
        """
        self._log("Double Tap element " + ranorexpath + " at location " + location + ".")
        if location == "":
            location = "Center"
        location = self._normalizeLocation(location)
        exec("Ranorex.Unknown(ranorexpath).DoubleTap(" + location + ")")

    def long_touch(self, ranorexpath, location = "Center", duration = "Ranorex.Mouse.DefaultMoveTime"):
        """ Performs a long touch on a mobile element.

        :param ranorexpath: RanoreXPath of the element that should get long touched.
        :param location: The location where the element should be touched. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight
        :param duration: Duration of the long touch event in ms.

        Example:
        | `Long Touch` | /mobileapp[@title='com.dropbox.android']//button[@accessiblename='Enter'] |  |  |
        | `Long Touch` | /mobileapp[@title='com.dropbox.android']//button[@accessiblename='Enter'] | LowerLeft |  |
        | `Long Touch` | /mobileapp[@title='com.dropbox.android']//button[@accessiblename='Enter'] | UpperRight | 3000 |
        """
        self._log("Long Touching element " + ranorexpath + " at location " + location + " for " + duration + "ms.")
        if location == "":
            location = "Center"
        if duration == "":
            duration = "Ranorex.Mouse.DefaultMoveTime"
        location = self._normalizeLocation(location)
        exec("Ranorex.Unknown(ranorexpath).LongTouch(" + location + ", " + duration + ")")

    def _moveTouchToElement(self, ranorexpath, location, duration):
        exec("Ranorex.Unknown(ranorexpath).MoveTo(" + location + ", " + duration + ")")

    def touch_start(self, ranorexpath, location = "Center"):
        """ Starts a touch event on a mobile element.

        :param ranorexpath: RanoreXPath of the element where the touch start event should be executed.
        :param location: The location where the element should be touched. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight

        Example:
        | `Touch Start` | /mobileapp[@title='com.dropbox.android']//container/androidelement/container[@containertype='Frame']/androidelement/container[9]/text | LowerLeft |
        """
        self._log("Touch Start on element" + ranorexpath + " at location " + location + ".")
        if location == "":
            location = "Center"
        location = self._normalizeLocation(location)
        exec("Ranorex.Unknown(ranorexpath).TouchStart(" + location + ")")

    def touch_move(self, ranorexpath, location = "Center", duration = "Ranorex.Mouse.DefaultMoveTime"):
        """ Moves to a specific element.

        Must be used after a touch start keyword and can be used to make complex gestures, like swiping.

        :param ranorexpath: RanoreXPath of the element where the move action _ends_.
        :param location: The location where the movement should end. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight
        :param duration: Duration of the movement in ms.

        Example:
        | `Touch Move` | /mobileapp[@title='com.dropbox.android']//container/androidelement/container[@containertype='Frame']/androidelement/container[9]/text | Center | 1500 |
        """
        self._log("Touch Move to element "+ ranorexpath + " at location " + location + " for " + duration + "ms.")
        if location == "":
            location = "Center"
        if duration == "":
            duration = "Ranorex.Mouse.DefaultMoveTime"
        location = self._normalizeLocation(location)
        self._moveTouchToElement(ranorexpath, location, duration)

    def touch_end(self, ranorexpath, location = "Center", duration = "Ranorex.Mouse.DefaultMoveTime"):
        """ This keyword lets a touch action end.

        It is normally preceded by a touch start or a touch start and several touch move events. It includes a move to the desired end location.

        :param location: The location where the touch should end. Possible values: Center, CenterLeft, CenterRight, LowerCenter, LowerRight, LowerLeft, UpperCenter, UpperLeft, UpperRight
        :param ranorexpath: RanoreXPath of the element on which the touch ends.
        :param duration: Duration of the movement to the end location in ms.

        Example:
        | `Touch End` | /mobileapp[@title='com.dropbox.android']//container[@containertype='Linear']/container[@containertype='Frame']/container[@containertype='Frame']/androidelement[1]/container[@containertype='Frame']/androidelement/container[@containertype='Frame']/container[3]/text | CenterRight | 50 |
        """
        self._log("Touch End on element " + ranorexpath + " at location " + location + " for " + duration + "ms.")
        if location == "":
            location = "Center"
        if duration == "":
            duration = "Ranorex.Mouse.DefaultMoveTime"
        location = self._normalizeLocation(location)
        self._moveTouchToElement(ranorexpath, location, duration)
        exec("Ranorex.Unknown(ranorexpath).TouchEnd(" + location + ")")

    def drag_and_drop(self, ranorexpath_src, ranorexpath_dest, location_src = "Center", location_dest = "Center", duration_src = "Ranorex.Mouse.DefaultMoveTime", duration_dest = "Ranorex.Mouse.DefaultMoveTime"):
        """ Performs a drag-and-drop action.

        This keyword is an action consisting of a mouse down, a mouse movement, and a mouse up. It is usually used to perform drag-and-drop actions, but also other actions (like selecting many cells in a table etc.) can be performed using this keyword.

        :param ranorexpath_src: RanoreXPath of the source element.
        :param ranorexpath_dest: RanoreXPath of the destination element.
        :param location_src: Where on the source element the mouse down action should happen.
        :param location_dest: Where on the destination element the mouse up action should happen.
        :param duration_src: How long it takes the mouse cursor to go to the surce element.
        :param duration_dest: How long it takes the mouse cursor to move to the destination element.

        Example:
        | `Drag and Drop` | form[@title='Desktop']/element[@class='ShellTabWindowClass']//element[@instance='1']/container[@caption='ShellView']/?/?/list/listitem[@automationid='0'] | /form[@title='Desktop']/element[@class='ShellTabWindowClass']//element[@instance='1']/container[@caption='ShellView']/?/?/list/listitem[@automationid='4'] |  |  |  |  |
        | `Drag and Drop` | form[@title='Desktop']/element[@class='ShellTabWindowClass']//element[@instance='1']/container[@caption='ShellView']/?/?/list/listitem[@automationid='0'] | /form[@title='Desktop']/element[@class='ShellTabWindowClass']//element[@instance='1']/container[@caption='ShellView']/?/?/list/listitem[@automationid='4'] | CenterLeft | UpperRight | 300 | 1500 |
        """
        self._log("Drag and drop from element " +  ranorexpath_src + " at location " + location_src + " to element " + ranorexpath_dest + " at location " + location_dest + ", taking " + duration_dest + "ms.")
        if location_src == "":
            location_src = "Center"
        if location_dest == "":
            location_dest = "Center"
        if duration_src == "":
            duration_src = "Ranorex.Mouse.DefaultMoveTime"
        if duration_dest == "":
            duration_dest = "Ranorex.Mouse.DefaultMoveTime"
        self.mouse_down(ranorexpath_src, location_src, "Left", duration_src)
        self.mouse_up(ranorexpath_dest, "Left", location_dest, duration_dest)

    def save_screenshot(self, ranorexpath, name, path):
        """ Saves a screenshot to the given location.

        :param ranorexpath: RanoreXPath of the element that a screenshot is taken of.
        :param name: Name of the image file. Should end with ".png"
        :param path: Path to the directory where the screenshot should be saved.

        Example:
        | `Save Screenshot` | /form[@controlname='RxMainFrame']//picture[@controlname='RxStudioLogo'] | logo.png | C:\\Users\\user\\Documents |
        """
        self._log("Capturing a screenshot of element " + ranorexpath + " and saving it to " + path + ".")
        image = Ranorex.Imaging.CaptureImageAuto(ranorexpath)
        pathName = path + "\\" + name
        image.Save(pathName, System.Drawing.Imaging.ImageFormat.Png)

    def report_screenshot(self, ranorexpath):
        """ Places a screenshot into the Robot test log file.

        The actual screenshot file is put into the working directory where the report and log files will also be generated by Robot.

        :param ranorexpath: RanoreXPath of the element that a screenshot is taken of.

        Example:
        | `Report Screenshot` | /form[@controlname='RxMainFrame']//picture[@controlname='RxStudioLogo'] |
        """
        self._log("Logging screenshot of element " + ranorexpath + ".")
        image = Ranorex.Imaging.CaptureImageAuto(ranorexpath)
        imgName = str(time.time()) + ".png"
        image.Save(imgName, System.Drawing.Imaging.ImageFormat.Png)
        htmlString = '<img src = "' + imgName + '">'
        logger.write(htmlString, self._logLevel, html = True)

    def invoke_action(self, ranorexpath, funcName, arguments = ""):
        self._log("Invoking action " + str(funcName) + " on element " + ranorexpath +  " with these arguments: " + arguments)
        if arguments == "":
            Ranorex.Unknown(ranorexpath).Element.InvokeActionWithText(funcName)
        else:
            Ranorex.Unknown(ranorexpath).Element.InvokeActionWithText(funcName, arguments)
