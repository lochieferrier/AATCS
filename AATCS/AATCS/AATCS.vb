'Last edited by Lochie Ferrier (Canberra Grammar School) on 29/05/2014
'Please contact lochieferrier@gmail.com for suppourt technical documentation
'Note that this program is experimental and as such may not direct aircraft in an appropriate manner in all situations
'Do not use for any real world air traffic application as the program is far too inefficient and unreliable for real use

Imports System.Speech.Synthesis
Public Class AATCS

    'Declare the top level variables.
    Dim vocaliseCommand As Boolean
    Dim printCommand As Boolean
    Dim numberOfAircraft As Integer
    Dim commandString As String
    Dim sector As Integer
    Dim executionCounter As Integer

    'Declare the Aircraft record structure.
    Public Structure Aircraft
        Public x As Integer
        Public y As Integer
        Public tgtHdg As Integer
        Public hdg As Integer
        Public tgtAltitude As Integer
        Public altitude As Integer
        Public target As Target
        Public callSign As String
        Public goAroundStatus As Boolean
    End Structure

    'Declare the Target record structure
    Public Structure Target
        Public x As Integer
        Public y As Integer
        Public approachHdg As Integer
        Public name As String
    End Structure


    Private Sub AATCS()

        'Note that aircraftArray is not initialized with a length
        Dim aircraftArray() As Aircraft
        Dim aircraftInArray As Boolean
        'The dynamic array is resized by ReadInAircraftFromFile which allows us to use GetUpperBound
        aircraftArray = ReadInAircraftFromFile(aircraftInArray)

        'Complete operations for the first execution (when the program starts up
        If executionCounter = 0 Then

            printCommand = False
            vocaliseCommand = False
            sector = 1
            KeyPreview = True

            lstOutput.Items.Add("Sector = " & sector)
            lstOutput.Items.Add("Vocalisation of commands = " & vocaliseCommand)
            lstOutput.Items.Add("Printing of commands = " & printCommand)
            lstOutput.Items.Add("Interval =  " & Timer.Interval & " ms")

            If aircraftInArray = True Then
                numberOfAircraft = aircraftArray.Length
                lstOutput.Items.Add("Initialized with " & aircraftArray.GetUpperBound(0) - 1 & " aircraft in array")
            Else
                lstOutput.Items.Add("Initialized with 0 aircraft in array")
                lstOutput.Items.Add("Add aircraft using the NUMA XXX command")
            End If
            lstOutput.Items.Add("")
            lstOutput.Items.Add("Type 'HELP' into the box below and click enter to view help")
            lstOutput.Items.Add("Or press F1 on your keyboard at any time to view help")
            lstOutput.Items.Add("")
        End If

        If aircraftInArray Then
            'Complete operations on aircraftArray by reference
            RouteAircraft(aircraftArray)
            AvoidCollisions(aircraftArray)
            'Use aircraftArray to generate commands
            IssueCommands(vocaliseCommand, aircraftArray)
        End If

        'Simulate and save aircraft movement, delete aircraft and create aircraft as necessary
        SimulateAndSaveAircraftMovementDeletionAndCreation(numberOfAircraft, aircraftArray, aircraftInArray)

        If aircraftInArray Then
            'Display aircraft and targets on screen
            DisplayAircraftAndTargets(sector, aircraftArray)
        End If

        executionCounter = executionCounter + 1

    End Sub

    Private Function ReadInAircraftFromFile(ByRef aircraftInArray As Boolean) As Array
        'This function reads in aircraft from the file "Aircraft.txt"
        Dim inputString As String
        Dim i As Integer
        Dim j As Integer
        Dim aircraftArray() As Aircraft
        aircraftInArray = False
        'Set initial values
        inputString = "9999"
        i = 0

        'Open file for reading
        FileSystem.FileOpen(1, Application.StartupPath & "\Data\Aircraft.txt", OpenMode.Input)

        'Check that the file is not empty
        If FileSystem.FileLen(Application.StartupPath & "\Data\Aircraft.txt") > 0 Then

            'Read in a line
            FileSystem.Input(1, inputString)

            'Iterate whilst there is only a reasonable number of elements or the EOF (9999) has not been reached
            While i <= 999 And inputString <> "9999"

                'As each line in the file consists of record fields seperated by commas, an inner loop is needed to read in each field
                j = 1
                Dim aircraft As Aircraft

                For j = 1 To 11

                    'This conditional statement ensures that the correct amount of lines are read in, taking into account the priming read
                    If (i <> 0 And j <> 1) Or (i = 0 And j <> 1) Then
                        FileSystem.Input(1, inputString)
                    End If
                    Select Case j
                        Case 1
                            aircraft.x = CInt(inputString)
                        Case 2
                            aircraft.y = CInt(inputString)
                        Case 3
                            aircraft.tgtHdg = CInt(inputString)
                        Case 4
                            aircraft.hdg = CInt(inputString)
                        Case 5
                            aircraft.tgtAltitude = CInt(inputString)
                        Case 6
                            aircraft.altitude = CInt(inputString)
                        Case 7
                            aircraft.target.x = CInt(inputString)
                        Case 8
                            aircraft.target.y = CInt(inputString)
                        Case 9
                            aircraft.target.approachHdg = CInt(inputString)
                        Case 10
                            aircraft.target.name = inputString
                        Case 11
                            aircraft.callSign = inputString
                    End Select
                Next

                FileSystem.Input(1, inputString)

                'Set the value of goAroundStatus by comparing strings
                If inputString = "0000T" Then
                    aircraft.goAroundStatus = True
                Else
                    aircraft.goAroundStatus = False
                End If

                'ReDim aircraftArray preserving all existing elements
                'This is what "resizes(" the array")
                ReDim Preserve aircraftArray(0 To i)
                'Store the newly read in aircraft in aircraftArray
                aircraftArray(i) = aircraft

                FileSystem.Input(1, inputString)

                i = i + 1
                aircraftInArray = True
            End While

        End If

        'It is vital that the file is closed as it is being accessed every second
        FileSystem.FileClose(1)
        Return aircraftArray

    End Function

    Private Sub RouteAircraft(ByRef aircraftArray() As Aircraft)
        'Iterate through the aircraft in aircraftArray, directing each towards its target
        For i = 0 To aircraftArray.GetUpperBound(0)

            Dim aircraft As Aircraft
            aircraft = aircraftArray(i)

            If aircraft.goAroundStatus = True Then
                If FindHorizontalDistanceToTarget(aircraft) > 15000 Then
                    aircraft.goAroundStatus = False
                End If
                SetTargetHeadingForAGoAroundManoueuvre(aircraft)
                SetTargetAltitudeForAGoAroundManoueuvre(aircraft)
            End If
            If aircraft.goAroundStatus = False And CheckIfAircraftIsOnApproachPath(aircraft) = False Then
                SetTargetAltitudeToApproachPath(aircraft)
                SetTargetHeadingToInterceptApproachPath(aircraft)
                Dim dHdg As Integer
                dHdg = FindDifferenceBetweenTwoHeadings(aircraft.hdg, aircraft.target.approachHdg)
                If FindHorizontalDistanceToTarget(aircraft) < 1000 And (FindHorizontalDistanceFromApproachPath(aircraft) > 500 Or FindVerticalDistanceFromApproachPath(aircraft) > 2000 Or dHdg > 40) Then
                    SetTargetHeadingForAGoAroundManoueuvre(aircraft)
                    SetTargetAltitudeForAGoAroundManoueuvre(aircraft)
                    aircraft.goAroundStatus = True
                End If
            End If
            aircraftArray(i) = aircraft
        Next
    End Sub

    Private Function CheckIfAircraftIsOnApproachPath(ByRef aircraft As Aircraft) As Boolean
        'Return a boolean indicating whether the aircraft is on the correct approach path
        Dim onPath As Boolean
        Dim dHdg As Integer
        dHdg = System.Math.Abs(aircraft.target.approachHdg - aircraft.hdg)
        'Use a conditional statement and special functions to determine whether the aircraft is on the approach path within a reasonable margin
        If FindVerticalDistanceFromApproachPath(aircraft) < 100 And FindHorizontalDistanceFromApproachPath(aircraft) < 1000 And dHdg < 5 Then
            onPath = True
        Else
            onPath = False
        End If

        Return onPath
    End Function

    Private Function FindVerticalDistanceFromApproachPath(ByVal aircraft As Aircraft) As Integer
        'Find the aircraft's vertical distance from its correct approach path
        Dim distance As Integer

        distance = FindHorizontalDistanceToTarget(aircraft)

        'Use the gradient of the approach path to calculate the height of the approach path at the current distance from the target and therefore calculate the vertical distance
        distance = System.Math.Abs(aircraft.altitude - distance * (4 / 73))

        Return distance
    End Function

    Private Function FindHorizontalDistanceFromApproachPath(ByVal aircraft As Aircraft) As Integer
        Dim angleToAxis As Integer
        Dim gradient As Decimal
        Dim constant As Integer
        Dim distance As Integer

        'Approach and other headings are stored as bearings so to find the angle to the x axis we need to subtract 90
        angleToAxis = aircraft.target.approachHdg
        'Use this angle to find the gradient
        gradient = System.Math.Tan(angleToAxis / (180 / System.Math.PI))
        'Find the intercept or constant of the approach path straight line equation
        constant = -gradient * (aircraft.target.x) + aircraft.target.y

        'Set distance to be the distance between the aircraft and the approach path using the System.Math.Abs function to return a positive value.
        distance = System.Math.Abs(-gradient * aircraft.x + aircraft.y - constant) / (((-gradient) ^ 2) + 1) ^ (1 / 2)

        Return distance
    End Function

    Private Sub SetTargetHeadingToInterceptApproachPath(ByRef aircraft As Aircraft)
        Dim proposedHdg As Integer
        Dim tempAircraft As Aircraft
        Dim minValue As Integer
        Dim count As Integer
        Dim value As Integer
        Dim optimalHdg As Integer
        Dim dx As Integer
        Dim dy As Integer

        proposedHdg = aircraft.hdg - 5
        optimalHdg = proposedHdg

        'Minimise the sum of the distance between the aircraft and target, distance between the aircraft and the approach path and the difference between the aircraft's heading and the approach heading.
        'This is of course a sub optimal equation and many improvements can be made to reduce go around rates and increase fuel efficiency.
        'However, these all require complicated mathematics that is not easy to implement and as such this equation was chosen for the first implementation.
        dx = (146 * System.Math.Cos((90 - proposedHdg) / (180 / System.Math.PI)))
        dy = -(146 * System.Math.Sin((90 - proposedHdg) / (180 / System.Math.PI)))

        tempAircraft = aircraft
        tempAircraft.x = tempAircraft.x + dx
        tempAircraft.y = tempAircraft.y + dy
        minValue = 10 * FindHorizontalDistanceToTarget(tempAircraft) + FindHorizontalDistanceFromApproachPath(tempAircraft) + FindDifferenceBetweenTwoHeadings(proposedHdg, aircraft.target.approachHdg)

        'Then iterate through all other available angles and find the correct path that maximizes the value
        For count = 0 To 10

            proposedHdg = proposedHdg + count
            dx = (146 * System.Math.Cos((90 - proposedHdg) / (180 / System.Math.PI)))
            dy = -(146 * System.Math.Sin((90 - proposedHdg) / (180 / System.Math.PI)))
            tempAircraft = aircraft
            tempAircraft.x = tempAircraft.x + dx
            tempAircraft.y = tempAircraft.y + dy
            value = 10 * FindHorizontalDistanceToTarget(tempAircraft) + FindHorizontalDistanceFromApproachPath(tempAircraft) + FindDifferenceBetweenTwoHeadings(proposedHdg, aircraft.target.approachHdg)

            If value < minValue Then
                optimalHdg = proposedHdg
            End If

            optimalHdg = formatHeading(optimalHdg)

        Next
        optimalHdg = formatHeading(optimalHdg)
        aircraft.tgtHdg = optimalHdg

    End Sub

    Private Sub SetTargetAltitudeToApproachPath(ByRef aircraft As Aircraft)
        'Set the aircraft target altitude to be the height of the approach path at the current distance
        aircraft.tgtAltitude = (4 / 73) * (FindHorizontalDistanceToTarget(aircraft))
    End Sub

    Private Function FindHorizontalDistanceToTarget(ByVal aircraft As Aircraft) As Integer
        'Find the distance between the aircraft and its target using Pythagoras' Theorem
        Dim distance As Integer
        Dim dx As Integer
        Dim dy As Integer
        distance = 0
        dx = aircraft.target.x - aircraft.x
        dy = aircraft.target.y - aircraft.y
        distance = System.Math.Sqrt(System.Math.Pow(dx, 2) + System.Math.Pow(dy, 2))
        Return distance
    End Function

    Private Sub SetTargetHeadingForAGoAroundManoueuvre(ByRef aircraft As Aircraft)
        'Set target heading for a go around manoueuvre, i.e. turn the aircraft around
        aircraft.tgtHdg = aircraft.target.approachHdg - 180
        aircraft.tgtHdg = FormatHeading(aircraft.tgtHdg)

    End Sub

    Private Sub SetTargetAltitudeForAGoAroundManoueuvre(ByRef aircraft As Aircraft)
        'Set the target altitude to 1000 (high enough to get out of the way of other approaching traffic)
        aircraft.tgtAltitude = 1000
    End Sub

    Private Sub AvoidCollisions(ByRef aircraftArray() As Aircraft)
        Dim count As Integer
        'Check that aircraftArraycontains multiple aircraft, as if it only contains one aircraft then there is no need to route aircraft to avoid collisions
        If aircraftArray.GetUpperBound(0) > 0 Then

            For count = 0 To aircraftArray.GetUpperBound(0) - 1

                'Create two aircraft, check for a risk of a collision between them and attempt to resolve
                Dim aircraftOne As Aircraft
                Dim aircraftTwo As Aircraft

                aircraftOne = aircraftArray(count)
                aircraftTwo = aircraftArray(count + 1)

                'If the aircraft are too close horizontally or vertically, divert them

                If FindHorizontalDistanceBetweenAircraft(aircraftOne, aircraftTwo) < 1000 Then
                    DivertAircraftHorizontally(aircraftOne, aircraftTwo)
                End If

                If FindVerticalDistanceBetweenAircraft(aircraftOne, aircraftTwo) < 100 Then
                    DivertAircraftVertically(aircraftOne, aircraftTwo)
                End If
                aircraftArray(count) = aircraftOne
                aircraftArray(count + 1) = aircraftTwo
            Next
        End If
    End Sub

    Private Function FindHorizontalDistanceBetweenAircraft(ByVal aircraftOne As Aircraft, ByVal aircraftTwo As Aircraft) As Integer
        'Find the horizontal distance between two aircraft using Pythagoras' Theorem
        Dim distance As Integer
        distance = System.Math.Sqrt((System.Math.Pow((aircraftOne.x - aircraftTwo.x), 2) + System.Math.Pow((aircraftOne.y - aircraftTwo.y), 2)))
        Return distance
    End Function

    Private Function FindVerticalDistanceBetweenAircraft(ByVal aircraftOne As Aircraft, ByVal aircraftTwo As Aircraft) As Integer
        'Find the vertical distance between aircraft using Pythagoras' Theorem
        Dim distance As Integer
        distance = System.Math.Abs(aircraftOne.altitude - aircraftTwo.altitude)
        Return distance
    End Function

    Private Sub DivertAircraftVertically(ByRef aircraftOne As Aircraft, ByRef aircraftTwo As Aircraft)
        'NOTE THAT THIS CODE HAS BEEN CHANGED POST SUBMISSION FOR THE PURPOSES OF TESTING

        'Seperate or divert aircraft vertically by directing the top aircraft to climb and the bottom aircraft to descend
        If aircraftOne.altitude >= aircraftTwo.altitude Then

            aircraftOne.tgtAltitude = aircraftOne.altitude + 1000
            aircraftTwo.tgtAltitude = aircraftTwo.altitude - 1000

            'Ensures that aircraftOne does not descend to a dangerous altitude
            If aircraftTwo.altitude - 1000 < 0 Then
                aircraftTwo.tgtAltitude = aircraftTwo.altitude
            End If

        Else

            aircraftTwo.tgtAltitude = aircraftTwo.altitude + 1000
            aircraftOne.tgtAltitude = aircraftOne.altitude - 1000

            'Ensures that aircraftOne does not descend to a dangerous altitude
            If aircraftOne.altitude - 1000 < 0 Then
                aircraftOne.tgtAltitude = aircraftOne.altitude
            End If

        End If
    End Sub

    Private Sub DivertAircraftHorizontally(ByRef aircraftOne As Aircraft, ByRef aircraftTwo As Aircraft)

        'Divert aircraft horizontally by changing course by 45 degrees multiplied by a random factor
        aircraftOne.tgtHdg = aircraftOne.hdg - Int(45 * ((Rnd() * 1.5) + 1))
        aircraftTwo.tgtHdg = aircraftTwo.hdg + Int(45 * ((Rnd() * 1.5) + 1))

        'Heading correction
        aircraftOne.tgtHdg = formatHeading(aircraftOne.tgtHdg)
        aircraftTwo.tgtHdg = formatHeading(aircraftTwo.tgtHdg)

    End Sub

    Private Sub IssueCommands(ByVal vocaliseCommand As Boolean, ByRef aircraftArray() As Aircraft)

        Dim aircraftFromFileArray() As Aircraft
        Dim count As Integer
        Dim aircraft As Aircraft
        Dim aircraftFromFile As Aircraft
        Dim aircraftInArray As Boolean
        'Read in aircraft from the file and compare with the aircraft in aircraftArray to determine whether there has been any commands given
        aircraftFromFileArray = ReadInAircraftFromFile(aircraftInArray)
        If aircraftInArray Then
            For count = 0 To aircraftFromFileArray.GetUpperBound(0)

                aircraft = aircraftArray(count)
                aircraftFromFile = aircraftFromFileArray(count)

                'Check whether the aircraft has changed in state (for example, the target heading has changed)
                If CheckForChangeInAircraftState(aircraft, aircraftFromFile) Then

                    'Generate, vocalise (if desired), print and save the command
                    Dim commandString As String
                    commandString = GenerateCommand(aircraft, aircraftFromFile)
                    If vocaliseCommand Then
                        VocaliseCommandUsingSystemSpeech(commandString)
                    End If
                    If printCommand Then
                        PrintCommandToListBox(commandString)
                    End If
                    SaveCommandToFile(commandString)

                End If
            Next
        End If
    End Sub

    Private Function CheckForChangeInAircraftState(ByVal aircraft As Aircraft, ByVal aircraftFromFile As Aircraft) As Boolean
        'Check whether any of the aircraft's target values (as in tgtAltitude or tgtHdg) have changed and return a boolean indicating whether this is the case
        Dim hasChanged As Boolean
        If aircraft.tgtAltitude <> aircraftFromFile.tgtAltitude Or aircraft.tgtHdg <> aircraftFromFile.tgtHdg Then
            hasChanged = True
        Else
            hasChanged = False
        End If
        Return hasChanged
    End Function

    Private Function GenerateCommand(ByVal aircraft As Aircraft, ByVal aircraftFromFile As Aircraft)
        'The purpose of this function is to generate a command if an aircraft has changed state.
        Dim commandString As String

        'Any command is prefixed by an aircraft's callsign for identification
        commandString = aircraft.callSign

        'Conditional statements are used to determine what command needs to be generated, for some aircraft a command
        If aircraft.tgtAltitude > aircraftFromFile.tgtAltitude Then
            commandString = commandString & " CLIMB " & aircraft.tgtAltitude
        End If

        If aircraft.tgtAltitude < aircraftFromFile.tgtAltitude Then
            commandString = commandString & " DESCEND " & aircraft.tgtAltitude
        End If

        If aircraft.tgtHdg <> aircraftFromFile.tgtHdg Then
            commandString = commandString & " HEADING " & aircraft.tgtHdg
        End If

        Return commandString

    End Function

    Private Sub VocaliseCommandUsingSystemSpeech(ByVal commandString As String)
        'The System.Speech class is used here to vocalise each command individually
        'Note that the commands can be vocalised using SpeakAsync however this usually causes a backup
        Dim speaker As New SpeechSynthesizer()
        speaker.Rate = 4
        speaker.Volume = 100
        speaker.Speak(commandString)

    End Sub

    Private Sub PrintCommandToListBox(ByVal commandString As String)
        'Print the command to lstOutput
        lstOutput.Items.Add(commandString)
    End Sub

    Private Sub SaveCommandToFile(ByVal commandString As String)
        'Open Commands.txt for appending
        FileSystem.FileOpen(1, Application.StartupPath & "\Data\Commands.txt", OpenMode.Append)
        'Add a UTC timestamp to the commandString
        commandString = commandString & " UTC: " & Date.UtcNow.Date & " " & Date.UtcNow.Hour & "h " & Date.UtcNow.Minute & "m " & Date.UtcNow.Second & "s "
        FileSystem.PrintLine(1, commandString)
        FileSystem.FileClose(1)
    End Sub

    Private Sub SimulateAndSaveAircraftMovementDeletionAndCreation(ByVal numberOfAircraft As Integer, ByRef aircraftArray() As Aircraft, ByRef aircraftInArray As Boolean)

        Dim count As Integer
        Dim aircraft As Aircraft
        If aircraftInArray Then
            'Iterate through all aircraft, simulating their motion through space
            For count = 0 To aircraftArray.GetUpperBound(0)
                aircraft = aircraftArray(count)
                SimulateHeadingChange(aircraft)
                SimulateHorizontalMovement(aircraft)
                SimulateAltitudeChange(aircraft)
                aircraftArray(count) = aircraft
            Next

            'Delete unnecessary aircraft
            DeleteAircraftIfTheyHaveReachedTheirTarget(aircraftArray)
            DeleteAircraftIfTheyAreOutOfBounds(aircraftArray)
        End If
        'Determine whether new aircraft are needed and if they are, create them
        If aircraftInArray = False Then
            If numberOfAircraft > 0 Then
                CreateNewAircraft(aircraftArray, numberOfAircraft, aircraftInArray)
                aircraftInArray = True
            End If
        ElseIf aircraftArray.GetUpperBound(0) + 1 < numberOfAircraft Then
            CreateNewAircraft(aircraftArray, numberOfAircraft, aircraftInArray)
            aircraftInArray = True
        End If

        'Save aircraftArray to file
        If aircraftInArray = True Then
            SaveAircraftToFile(aircraftArray)
        End If


    End Sub

    Private Sub SimulateHeadingChange(ByRef aircraft As Aircraft)

        'Case for if the difference between the aircraft's target heading and current heading is greater than the maximum turn angle (+-5) for one time step
        If System.Math.Abs(aircraft.tgtHdg - aircraft.hdg) > 5 Then
            If aircraft.tgtHdg > aircraft.hdg Then
                aircraft.hdg = aircraft.hdg + 5
            Else
                aircraft.hdg = aircraft.hdg - 5
            End If
        End If

        'Case for if the difference between the aircraft's target heading and current heading is less than the maximum turn angle (+-5) for one time step
        If 0 < System.Math.Abs(aircraft.tgtHdg - aircraft.hdg) < 5 Then
            'Oscillation is avoided by directly setting value of aircraft.hdg to be aircraft.tgtHdg
            aircraft.hdg = aircraft.tgtHdg
        End If
    End Sub

    Private Sub SimulateHorizontalMovement(ByRef aircraft As Aircraft)

        'Use trigonometry to calculate the next position of the aircraft
        aircraft.x = aircraft.x + (146 * System.Math.Cos((90 - aircraft.hdg) / (180 / System.Math.PI)))
        aircraft.y = aircraft.y - (146 * System.Math.Sin((90 - aircraft.hdg) / (180 / System.Math.PI)))

    End Sub

    Private Sub SimulateAltitudeChange(ByRef aircraft As Aircraft)
        'Simulate a change in altitude using a climb rate of 23m/s and a descent rate of 11m/s
        If aircraft.tgtAltitude > aircraft.altitude Then
            If (aircraft.tgtAltitude - aircraft.altitude) > 23 Then
                aircraft.altitude = aircraft.altitude + 23
            End If
            If (aircraft.tgtAltitude - aircraft.altitude) <= 23 Then
                aircraft.altitude = aircraft.tgtAltitude
            End If
        End If

        If aircraft.tgtAltitude < aircraft.altitude Then
            If (aircraft.altitude - aircraft.tgtAltitude) > 11 Then
                aircraft.altitude = aircraft.altitude - 11
            End If
            If (aircraft.altitude - aircraft.tgtAltitude) <= 11 Then
                aircraft.altitude = aircraft.tgtAltitude
            End If
        End If

    End Sub

    Private Sub DeleteAircraftIfTheyHaveReachedTheirTarget(ByRef aircraftArray() As Aircraft)

        Dim i As Integer
        Dim aircraft As Aircraft
        i = 0
            While i <= aircraftArray.GetUpperBound(0)
            aircraft = aircraftArray(i)

            'Determine whether the aircraft has arrived at the target at the correct heading, position and altitude
            If FindHorizontalDistanceToTarget(aircraft) < 1000 And FindDifferenceBetweenTwoHeadings(aircraft.hdg, aircraft.tgtHdg) < 20 And aircraft.altitude < 1000 Then
                ShuffleArray(i, aircraftArray)
                If aircraftArray.GetUpperBound(0) > 0 Then
                    ReDim Preserve aircraftArray(0 To aircraftArray.GetUpperBound(0) - 1)
                Else
                    ReDim aircraftArray(0 To 0)
                End If
                lstOutput.Items.Add("Aircraft " & aircraft.callSign & " reached target " & aircraft.target.name)
            Else
                i = i + 1
            End If
            i = i + 1
        End While
    End Sub

    Private Sub DeleteAircraftIfTheyAreOutOfBounds(ByRef aircraftArray() As Aircraft)

        Dim j As Integer
        Dim aircraft As Aircraft

        j = 0

        While j <= aircraftArray.GetUpperBound(0)

            aircraft = aircraftArray(j)
            'Determine whether the aircraft is within the displayable and saveable region
            If aircraft.x < 0 Or aircraft.y < 0 Or aircraft.x > 90000 Or aircraft.y > 90000 Or aircraft.altitude < 0 Then
                ShuffleArray(j, aircraftArray)
                If aircraftArray.GetUpperBound(0) > 0 Then
                    ReDim Preserve aircraftArray(0 To aircraftArray.GetUpperBound(0) - 1)
                Else
                    ReDim aircraftArray(0 To 0)
                End If

            Else

                j = j + 1

            End If
        End While
    End Sub

    Private Sub ShuffleArray(ByVal index As Integer, ByRef aircraftArray() As Aircraft)
        'Shuffles all the elements after the index back one element
        Dim count As Integer
        For count = index + 1 To aircraftArray.GetUpperBound(0)
            aircraftArray(count - 1) = aircraftArray(count)
        Next
    End Sub

    Private Sub CreateNewAircraft(ByRef aircraftArray() As Aircraft, ByVal numberOfAircraft As Integer, ByVal aircraftInArray As Boolean)
        'If the aircraftArray does not have any aircraft, create a blank aircraft which will be used for creating the first aircraft
        If aircraftInArray = False Then
            ReDim aircraftArray(0)
        End If
        If numberOfAircraft <= 999 Then
            If numberOfAircraft > aircraftArray.GetUpperBound(0) Then
                Dim targetsArray() As Target
                targetsArray = ReadTargetsFromFile()
                Do
                    Dim aircraft As Aircraft
                    'Randomly generate each field of the aircraft, using the Randomize function to ensure that aircraft are not clustered
                    aircraft.callSign = GenerateRandomCallSign(aircraftArray)
                    Randomize()
                    aircraft.x = Int((Rnd() * 80000) + 10000)
                    aircraft.y = Int((Rnd() * 80000) + 10000)
                    aircraft.altitude = Int((Rnd() * 3000) + 0)
                    aircraft.tgtAltitude = aircraft.altitude
                    aircraft.hdg = Int((Rnd() * 359) + 0)
                    aircraft.tgtHdg = aircraft.hdg
                    aircraft.target = targetsArray(CInt((Rnd() * (targetsArray.GetUpperBound(0)))))
                    If aircraftInArray = False Then
                        aircraftArray(0) = aircraft
                        aircraftInArray = True
                    Else
                        ReDim Preserve aircraftArray(0 To aircraftArray.GetUpperBound(0) + 1)
                        aircraftArray(aircraftArray.GetUpperBound(0)) = aircraft
                    End If
                Loop Until aircraftArray.GetUpperBound(0) + 1 = numberOfAircraft
            End If
        End If
    End Sub

    Private Function GenerateRandomCallSign(ByVal aircraftArray() As Aircraft) As String

        Dim unique As Boolean
        Dim randomCallSign As String
        Dim count As Integer
        Dim i As Integer
        Dim aircraft As Aircraft

        'Create random callsigns until a unique one is found
        Do
            unique = True
            'This prefix can be changed or be generated randomly
            randomCallSign = "QF"
            'Use a FOR loop to append 3 random digits on to the end of the call sign
            For i = 1 To 3
                randomCallSign = randomCallSign & CStr(CInt((Rnd() * 9) + 0))
            Next

            'This FOR loop checks whether the callsign is already in use
            For count = 0 To aircraftArray.GetUpperBound(0)
                aircraft = aircraftArray(count)
                If aircraft.callSign = randomCallSign Then
                    unique = False
                End If
            Next
        Loop Until unique = True
        Return randomCallSign
    End Function

    Private Function ReadTargetsFromFile() As Array
        Dim line As String
        Dim i As Integer
        Dim j As Integer
        Dim targetArray() As Target
        FileSystem.FileOpen(1, Application.StartupPath & "\Data\Targets.txt", OpenMode.Input)

        'Set initial values
        line = "9999"
        i = 0

        FileSystem.Input(1, line)

        'Read lines from the file until too many lines have been read in or the sentinel value of 9999 has been reached
        While i <= 999 And line <> "9999"
            j = 1
            Dim target As Target

            For j = 1 To 4

                'The purpose of this conditional is to ensure that the correct number of lines are read in, taking into account the priming read
                If (i <> 0 And j <> 1) Or (i = 0 And j <> 1) Then
                    FileSystem.Input(1, line)
                End If

                Select Case j
                    Case 1
                        target.x = CInt(line)
                    Case 2
                        target.y = CInt(line)
                    Case 3
                        target.approachHdg = CInt(line)
                    Case 4
                        target.name = line
                End Select

            Next j

            FileSystem.Input(1, line)
            ReDim Preserve targetArray(0 To i)
            targetArray(i) = target

            i = i + 1
        End While
        FileSystem.FileClose(1)

        Return targetArray

    End Function

    Private Sub SaveAircraftToFile(ByVal aircraftArray() As Aircraft)
        'Erase the contents of Aircraft.txt
        FileSystem.FileOpen(1, Application.StartupPath & "\Data\Aircraft.txt", OpenMode.Output)

        Dim count As Integer
        Dim aircraft As Aircraft
        Dim line As String
        For count = 0 To aircraftArray.GetUpperBound(0)

            line = ""
            aircraft = aircraftArray(count)
            'Write out each field on the line, seperating with commas and converting numbers to strings with leading zeroes
            line = line + ConvertIntegerToStringWithLeadingZeroes(aircraft.x, 5) + ","
            line = line + ConvertIntegerToStringWithLeadingZeroes(aircraft.y, 5) + ","
            line = line + ConvertIntegerToStringWithLeadingZeroes(aircraft.tgtHdg, 5) + ","
            line = line + ConvertIntegerToStringWithLeadingZeroes(aircraft.hdg, 5) + ","
            line = line + ConvertIntegerToStringWithLeadingZeroes(aircraft.tgtAltitude, 5) + ","
            line = line + ConvertIntegerToStringWithLeadingZeroes(aircraft.altitude, 5) + ","
            line = line + ConvertIntegerToStringWithLeadingZeroes(aircraft.target.x, 5) + ","
            line = line + ConvertIntegerToStringWithLeadingZeroes(aircraft.target.y, 5) + ","
            line = line + ConvertIntegerToStringWithLeadingZeroes(aircraft.target.approachHdg, 5) + ","
            line = line + aircraft.target.name + ","
            line = line + aircraft.callSign + ","
            If aircraft.goAroundStatus Then
                line = line + "0000T"
            Else
                line = line + "0000F"
            End If

            FileSystem.PrintLine(1, line)

        Next

        FileSystem.PrintLine(1, "9999")
        FileSystem.FileClose(1)

    End Sub

    Private Function ConvertIntegerToStringWithLeadingZeroes(ByVal number As Integer, ByVal stringLength As Integer) As String
        Dim numberString As String
        Dim difference As Integer
        numberString = CStr(number)
        difference = stringLength - numberString.Length
        If difference >= 1 Then
            'Use a FOR loop to add the required number of zeroes to the start of the string
            For count = 0 To difference - 1
                numberString = "0" + numberString
            Next
        End If
        Return numberString
    End Function

    Private Sub DisplayAircraftAndTargets(ByVal sector As Integer, ByVal aircraftArray() As Aircraft)

        'Reset the drawing surface
        SetUpDrawingSurface()
        'Draw the targets in the sector currently being displayed
        DrawTargets(sector)
        'Create aircraftArrayBySector as an empty dynamic 2D array
        Dim aircraftArrayBySector(,) As Aircraft
        'Load aircraftArrayBySector with the aircraft from aircraftArray
        aircraftArrayBySector = SortAircraftIntoSectors(aircraftArray)
        'Draw the aircraft in the current sector being displayed
        DrawAircraft(sector, aircraftArrayBySector)

    End Sub

    Private Sub SetUpDrawingSurface()
        'Create a new System.Drawing.Graphics object and clear the drawing canvas, replacing with black
        Dim g As System.Drawing.Graphics
        g = Me.CreateGraphics()
        g.Clear(Color.Black)
    End Sub

    Private Sub DrawTargets(ByVal sector As Integer)

        'Create a System.Drawing.Graphics object to use for drawing
        Dim g As System.Drawing.Graphics
        g = Me.CreateGraphics()

        Dim targetArray() As Target
        'Note that targetArrayBySector is created here with a length in the second dimension enough to store the maximum number of aircraft
        Dim targetArrayBySector(9, 999) As Target
        Dim count As Integer
        Dim target As Target

        'Read in the targets from Targets.txt
        targetArray = ReadTargetsFromFile()
        'Sort targetArray into targetArrayBySector
        targetArrayBySector = SortTargetsIntoSectors(targetArray)

        For count = 0 To 999
            target = targetArrayBySector(sector, count)
            'Here it is necessary to check whether the target object is valid as the array is usually full of empty target objects
            If target.name <> "" Then
                'This subroutine changes the target's coordinates by reference, though it does not affect the target stored in the file as the targets are not saved
                ConvertTargetCoordinatesToScreenCoordinates(sector, target)
                Dim pinkPen As New Drawing.Pen(Color.Pink, 2)
                'Draw an ellipse with at the target's (screen converted) coordinates
                g.DrawEllipse(pinkPen, target.x, target.y, 5, 5)
                Dim fontObj As Font
                fontObj = New System.Drawing.Font("Courier", 5, FontStyle.Regular)
                g.DrawString(target.name & vbCrLf & "HDG " & target.approachHdg, fontObj, Brushes.Pink, target.x, target.y + 10)
            End If

        Next
    End Sub

    Private Function SortTargetsIntoSectors(ByVal targetArray() As Target) As Array
        'Note that the arrays in the subroutine must be used as fixed length arrays as it is difficult to resize multidimensional arrays
        Dim count As Integer
        Dim targetArrayByLevel(2, 1000) As Target
        Dim targetArrayBySector(10, 1000) As Target
        Dim level As Integer
        Dim i As Integer
        Dim j As Integer

        'Sort target array by y coordinates using an insertion sort
        If targetArray.GetUpperBound(0) > 0 Then

            For i = 1 To targetArray.GetUpperBound(0)

                Dim tempTarget As Target
                tempTarget = targetArray(i)

                For j = i To 1 Step -1

                    If targetArray(j - 1).y > tempTarget.y Then
                        targetArray(j) = targetArray(j - 1)
                    Else
                        Exit For
                    End If

                Next j

                targetArray(j) = tempTarget

            Next i
        End If

        'Split up target array by y coordinates into three 'levels'
        Dim target As Target
        For count = 0 To targetArray.GetUpperBound(0)
            target = targetArray(count)
            If target.y >= 30000 Then
                If target.y >= 60000 Then
                    targetArrayByLevel(2, count) = target
                Else
                    targetArrayByLevel(1, count) = target
                End If
            Else
                targetArrayByLevel(0, count) = target
            End If
        Next

        'Use an insertion sort to sort the levels by x coordinates.

        For level = 0 To 2

            For i = 1 To 999

                Dim tempTarget As Target
                tempTarget = targetArrayByLevel(level, i)

                For j = i To 1 Step -1

                    If targetArrayByLevel(level, j - 1).x > tempTarget.x Then
                        targetArrayByLevel(level, j) = targetArrayByLevel(level, j - 1)
                    Else
                        Exit For
                    End If

                Next j

                targetArrayByLevel(level, j) = tempTarget

            Next i

        Next

        'Split each level into sectors.

        For level = 1 To 3
            For count = 0 To 999
                target = targetArrayByLevel(level - 1, count)
                If target.name <> "" Then
                    If target.x >= 30000 Then
                        If target.x >= 60000 Then
                            targetArrayBySector(level * 3, count) = target
                        Else
                            targetArrayBySector((level * 3) - 1, count) = target
                        End If
                    Else
                        targetArrayBySector((level * 3) - 2, count) = target
                    End If
                End If
            Next
        Next
        Return targetArrayBySector


    End Function

    Private Function SortAircraftIntoSectors(ByVal aircraftArray() As Aircraft) As Array
        'Note that the arrays in the subroutine must be used as fixed length arrays as it is difficult to resize multidimensional arrays
        Dim count As Integer
        Dim aircraftArrayByLevel(2, 1000) As Aircraft
        Dim aircraftArrayBySector(10, 1000) As Aircraft
        Dim level As Integer
        Dim i As Integer
        Dim j As Integer

        'Sort aircraft array by y coordinates using an insertion sort
        If aircraftArray.GetUpperBound(0) > 0 Then

            For i = 1 To aircraftArray.GetUpperBound(0)

                Dim tempAircraft As Aircraft
                tempAircraft = aircraftArray(i)

                For j = i To 1 Step -1

                    If aircraftArray(j - 1).y > tempAircraft.y Then
                        aircraftArray(j) = aircraftArray(j - 1)
                    Else
                        Exit For
                    End If

                Next j

                aircraftArray(j) = tempAircraft

            Next i
        End If
        'Split up aircraft array by y coordinates into three 'levels'
        Dim aircraft As Aircraft
        For count = 0 To aircraftArray.GetUpperBound(0)
            aircraft = aircraftArray(count)
            If aircraft.y >= 30000 Then
                If aircraft.y >= 60000 Then
                    aircraftArrayByLevel(2, count) = aircraft
                Else
                    aircraftArrayByLevel(1, count) = aircraft
                End If
            Else
                aircraftArrayByLevel(0, count) = aircraft
            End If
        Next

        'Use an insertion sort to sort the levels by x coordinates.

        For level = 0 To 2

            For i = 1 To 999

                Dim tempAircraft As Aircraft
                tempAircraft = aircraftArrayByLevel(level, i)

                For j = i To 1 Step -1

                    If aircraftArrayByLevel(level, j - 1).y > tempAircraft.y Then
                        aircraftArrayByLevel(level, j) = aircraftArrayByLevel(level, j - 1)
                    Else
                        Exit For
                    End If

                Next j

                aircraftArrayByLevel(level, j) = tempAircraft

            Next i

        Next

        'Split each level into sectors.

        For level = 1 To 3
            For count = 0 To 999
                aircraft = aircraftArrayByLevel(level - 1, count)
                If aircraft.callSign <> "" Then
                    If aircraft.x >= 30000 Then
                        If aircraft.x >= 60000 Then
                            aircraftArrayBySector(level * 3, count) = aircraft

                        Else
                            aircraftArrayBySector((level * 3) - 1, count) = aircraft

                        End If
                    Else
                        aircraftArrayBySector((level * 3) - 2, count) = aircraft

                    End If
                End If
            Next
        Next
        Return aircraftArrayBySector


    End Function
    Private Sub DrawAircraft(ByVal sector As Integer, ByVal aircraftArrayBySector(,) As Aircraft)
        'Create a System.Drawing.Graphics object to use for the drawing
        Dim g As System.Drawing.Graphics
        g = Me.CreateGraphics()

        Dim count As Integer
        Dim aircraft As Aircraft

        For count = 0 To 999
            aircraft = aircraftArrayBySector(sector, count)
            'Check whether the aircraft is valid by checking callsign as aircraftArrayBySector usually contains many blank elements
            If aircraft.callSign <> "" Then

                'This subroutine converts the aircraft coordinates by reference to screen drawing coordinates, though the aircraft in the file is unaffected as the aircraft are not saved
                ConvertAircraftCoordinatesToScreenCoordinates(sector, aircraft)
                Dim rectPen As New Drawing.Pen(Color.Green, 1)

                'Change the colour of the rectangle depending on the aircraft's goAroundStatus to assist with debugging and general use
                If aircraft.goAroundStatus = True Then
                    rectPen.Color = Color.Red
                End If

                'Draw a rectangle with at the aircraft's (screen converted) coordinates
                g.DrawRectangle(rectPen, aircraft.x, aircraft.y, 5, 5)
                Dim fontObj As Font
                fontObj = New System.Drawing.Font("Courier", 5, FontStyle.Regular)
                'Draw a label below the aircraft with basic information
                g.DrawString("ALT " & aircraft.altitude & vbCrLf & "HDG " & aircraft.hdg & vbCrLf & "SIGN " & aircraft.callSign & vbCrLf & "TGT " & aircraft.target.name, fontObj, Brushes.Orange, aircraft.x, aircraft.y + 10)

            End If
        Next

    End Sub

    Private Sub ConvertAircraftCoordinatesToScreenCoordinates(ByVal sector As Integer, ByRef aircraft As Aircraft)

        Dim xResolution As Integer
        Dim yResolution As Integer
        xResolution = 580
        yResolution = 400

        'These equations convert coordinates by taking into account the fact that sectors measure 30000 * 30000 and one sector is displayed on screen
        'If this is known then using the fixed x and y resolutions, screen coordinates can be found
        aircraft.x = CInt((aircraft.x Mod 30000) * (xResolution / 30000))
        aircraft.y = CInt((aircraft.y Mod 30000) * (yResolution / 30000))

    End Sub

    Private Sub ConvertTargetCoordinatesToScreenCoordinates(ByVal sector As Integer, ByRef target As Target)

        Dim xResolution As Integer
        Dim yResolution As Integer
        xResolution = 580
        yResolution = 400

        'These equations convert coordinates by taking into account the fact that sectors measure 30000 * 30000 and one sector is displayed on screen
        'If this is known then using the fixed x and y resolutions, screen coordinates can be found
        target.x = CInt((target.x Mod 30000) * (xResolution / 30000))
        target.y = CInt((target.y Mod 30000) * (yResolution / 30000))

    End Sub

    Private Sub ExecuteTextCommands(ByVal commandString As String, ByRef numberOfAircraft As Integer, ByRef vocaliseCommand As Boolean, ByRef sector As Integer, ByRef printCommand As Boolean)

        'Check whether the input string is valid before progressing further
        If commandString.Length() <= 9 Then
            'Determine what subroutine should handle the command by using the first 4 characters of the commandString
            Select Case commandString.Substring(0, 4)
                Case "SECT"
                    ChangeSector(sector, commandString)
                Case "VOXC"
                    TurnOnOrOffVocalisationOfCommands(vocaliseCommand, commandString)
                Case "NUMA"
                    ChangeNumberOfAircraft(numberOfAircraft, commandString)
                Case "PRNT"
                    TurnOnOrOffPrintingOfCommands(printCommand, commandString)
                Case "TIME"
                    ChangeTimeStep(commandString)
                Case "HELP"
                    DisplayHelp()
            End Select
        Else
            'Tell the user that an invalid command has been entered
            lstOutput.Items.Add("Invalid input")
        End If

    End Sub

    Private Sub ChangeSector(ByRef sector As Integer, ByRef commandString As String)

        'Check that the length of the command entered is correct
        If commandString.Length = 6 Then

            'Set sector to be the integer value of the 5th (or 6th from 0) character
            sector = CInt(CStr(commandString(5)))

            'Ensure that sector is not an invalid value
            If sector = 0 Then
                sector = 1
            End If

            lstOutput.Items.Add("Sector changed to: " & sector)
        Else
            'Tell the user than an invalid command has been entered
            lstOutput.Items.Add("Invalid input")
        End If

    End Sub

    Private Sub TurnOnOrOffVocalisationOfCommands(ByRef vocaliseCommand As Boolean, ByVal commandString As String)

        'Check that the length of the command entered is correct
        If commandString.Length = 6 Then
            'If the commandString entered is N (for no) then disable vocalisation of commands, otherwise enable vocalisation of commands
            If CStr(commandString(5)) = "N" Then
                vocaliseCommand = False
                lstOutput.Items.Add("Vocalisation of commands has been disabled")
            Else
                vocaliseCommand = True
                lstOutput.Items.Add("Vocalisation of commands has been enabled")
            End If
        Else
            'Tell the user than an invalid command has been entered
            lstOutput.Items.Add("Invalid input")
        End If

    End Sub

    Private Sub TurnOnOrOffPrintingOfCommands(ByRef printCommand As Boolean, ByVal commandString As String)

        'Check that the length of the command entered is correct
        If commandString.Length = 6 Then
            'If the commandString entered is N (for no) then disable printing of commands, otherwise enable printing of commands
            If CStr(commandString(5)) = "N" Then
                printCommand = False
                lstOutput.Items.Add("Printing of commands has been disabled")
            Else
                printCommand = True
                lstOutput.Items.Add("Printing of commands has been enabled")
            End If
        Else
            'Tell the user than an invalid command has been entered
            lstOutput.Items.Add("Invalid input")
        End If

    End Sub


    Private Sub ChangeNumberOfAircraft(ByRef numberOfAircraft As Integer, ByVal commandString As String)

        'Check that the length of the command entered by the user is correct (to ensure that leading zeroes have been included)
        If commandString.Length = 8 Then
            'Set numberOfAircraft to be the integer value of the 6th to 9th characters of commandString
            numberOfAircraft = CInt(commandString.Substring(5, 3))
            'If the number entered is too great set numberOfAircraft to be 999
            If numberOfAircraft > 999 Then
                numberOfAircraft = 999
            End If

            lstOutput.Items.Add("Number of aircraft changed to: " & numberOfAircraft)

        Else
            'Tell the user that an invalid command has been entered
            lstOutput.Items.Add("Invalid input")
        End If

    End Sub

    Private Sub ChangeTimeStep(ByVal commandString As String)

        'Check that the length of the command entered by the user is correct (to ensure that leading zeroes have been included)
        If commandString.Length = 9 Then
            'Set the interval
            Timer.Interval = CInt(commandString.Substring(5, 4))
            If Timer.Interval < 10 Then
                Timer.Interval = 10
            End If
            lstOutput.Items.Add("Time step changed to: " & Timer.Interval & " ms")
        Else
            'Tell the user that an invalid command has been entered
            lstOutput.Items.Add("Invalid input")
        End If

    End Sub

    Private Sub AATCSKeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        'If the key released was F1 then run the DisplayHelp subroutine
        If e.KeyData = Keys.F1 Then
            DisplayHelp()
        End If

    End Sub

    Private Sub DisplayHelp()
        'Open the User Guide pdf
        System.Diagnostics.Process.Start(Application.StartupPath & "\Help\User Guide.pdf")
        lstOutput.Items.Add("Help displayed")
    End Sub

    Private Function FormatHeading(ByVal hdg As Integer) As Integer

        'The purpose of this function is to ensure that no headings outside the domain of 0-359 are used

        'Check whether heading is negative and correct if necessary
        If hdg < 0 Then
            hdg = hdg + 359
        End If

        'Check whether heading is too large and correct if necessary
        If hdg > 359 Then
            hdg = hdg - 360
        End If

        Return hdg

    End Function

    Private Function FindDifferenceBetweenTwoHeadings(ByVal headingOne As Integer, ByVal headingTwo As Integer) As Integer
        Dim minDifference As Integer
        minDifference = 1000
        Dim dHdg As Integer
        dHdg = 0
        'All four methods of calculating the difference between the two headings to find the smallest (and therefore valid) difference between them
        'A much more complex approach would determine the quadrants of the two headings but it would take more lines of code and be less readable

        'These two conditionals are for the top case, so both headings cause an increase in y value if an aircraft follows them
        'Or more simply put, both headings point upwards
        If System.Math.Abs((360 - headingOne) + headingTwo) < minDifference Then
            dHdg = System.Math.Abs((360 - headingOne) + headingTwo)
            minDifference = dHdg
        End If
        If System.Math.Abs((360 - headingTwo) + headingOne) < minDifference Then
            dHdg = (360 - headingTwo) + headingOne
            minDifference = dHdg
        End If

        'These two conditionals are for the top case, so both headings cause a decrease in y value if an aircraft follows them
        'Or more simply put, both headings point downwards
        If System.Math.Abs(headingOne - headingTwo) < minDifference Then
            dHdg = headingOne - headingTwo
            minDifference = dHdg
        End If
        If System.Math.Abs(headingTwo - headingOne) < minDifference Then
            minDifference = dHdg
        End If

        Return dHdg

    End Function

    Private Sub btnEnter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnter.Click

        'If the Enter button is clicked then the ExecuteTextCommands subroutine is run
        ExecuteTextCommands(txtCommand.Text, numberOfAircraft, vocaliseCommand, sector, printCommand)

    End Sub

    Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTest.Click
        TestSubroutinesUsingDrivers()
    End Sub

    Private Sub TestSubroutinesUsingDrivers()

        'FindHorizontalDistanceFromApproachPath driver
        Dim aircraft As Aircraft
        aircraft.x = 50
        aircraft.y = 100
        aircraft.target.approachHdg = 225
        aircraft.target.x = 200
        aircraft.target.y = 200
        aircraft.altitude = 1000
        aircraft.tgtAltitude = 1000
        aircraft.hdg = 90
        aircraft.tgtHdg = 90
        aircraft.callSign = "FTEST"
        aircraft.target.name = "TEST0"
        aircraft.goAroundStatus = False

        lstOutput.Items.Add("FindHorizontalDistanceFromApproachPath returned:")
        lstOutput.Items.Add(FindHorizontalDistanceFromApproachPath(aircraft))

        FindHorizontalDistanceToTarget(driver)
        Dim aircraft As Aircraft
        aircraft.x = 200
        aircraft.y = 200
        aircraft.target.approachHdg = 180
        aircraft.target.x = 200
        aircraft.target.y = 200
        aircraft.altitude = 1000
        aircraft.tgtAltitude = 1000
        aircraft.hdg = 90
        aircraft.tgtHdg = 90
        aircraft.callSign = "FTEST"
        aircraft.target.name = "TEST0"
        aircraft.goAroundStatus = False

        lstOutput.Items.Add("FindHorizontalDistanceToTarget returned:")
        lstOutput.Items.Add(FindHorizontalDistanceToTarget(aircraft))

        'DivertAircraftVertically driver
        Dim aircraftOne As Aircraft
        aircraftOne.x = 0
        aircraftOne.y = 0
        aircraftOne.target.approachHdg = 180
        aircraftOne.target.x = 200
        aircraftOne.target.y = 200
        aircraftOne.altitude = 4500
        aircraftOne.tgtAltitude = 9999
        aircraftOne.hdg = 90
        aircraftOne.tgtHdg = 90
        aircraftOne.callSign = "TEST1"
        aircraftOne.target.name = "TEST0"
        aircraftOne.goAroundStatus = False

        Dim aircraftTwo As Aircraft
        aircraftTwo.x = 0
        aircraftTwo.y = 0
        aircraftTwo.target.approachHdg = 180
        aircraftTwo.target.x = 200
        aircraftTwo.target.y = 200
        aircraftTwo.altitude = 4250
        aircraftTwo.tgtAltitude = 9999
        aircraftTwo.hdg = 90
        aircraftTwo.tgtHdg = 90
        aircraftTwo.callSign = "TEST2"
        aircraftTwo.target.name = "TEST0"
        aircraftTwo.goAroundStatus = False

        DivertAircraftVertically(aircraftOne, aircraftTwo)
        lstOutput.Items.Add("Diverting the two aircraft vertically changed target values to:")
        lstOutput.Items.Add("aircraftOne tgtAltitude: " & aircraftOne.tgtAltitude)
        lstOutput.Items.Add("aircraftTwo tgtAltitude: " & aircraftTwo.tgtAltitude)

    End Sub

    
End Class