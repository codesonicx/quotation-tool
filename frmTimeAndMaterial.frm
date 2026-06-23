Option Explicit

' =========================================================
' Form configuration constants
' =========================================================
Private Const FORM_HEIGHT As Long = 420
Private Const FORM_WIDTH As Long = 580
Private Const FORM_TITLE As String = "Time and Material"

' Prevents recursive Change events when weeks/days update each other.
Private isRefreshing As Boolean

' =========================================================
' Form lifecycle
' =========================================================

Private Sub UserForm_Initialize()
    ConfigureForm
    LoadDefaultsFromDrivers
    RefreshTotals
End Sub

Private Sub UserForm_Activate()
    CenterFormOnExcel Me
End Sub

' =========================================================
' Button events
' =========================================================

Private Sub btnCreate_Click()
    HandleCreateAction
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

' =========================================================
' Reactive input events
' =========================================================

Private Sub txtStartDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UpdateDurationFromDates
End Sub

Private Sub txtFinalDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UpdateDurationFromDates
End Sub

Private Sub txtNumWeeks_Change()
    If isRefreshing Then Exit Sub

    isRefreshing = True
    Me.txtNumDays.Value = CStr(Round(ToNumber(Me.txtNumWeeks.Value) * 7, 1))
    UpdateFinalDateFromDays
    isRefreshing = False

    RefreshTotals
End Sub

Private Sub txtNumDays_Change()
    If isRefreshing Then Exit Sub

    isRefreshing = True
    Me.txtNumWeeks.Value = CStr(Round(ToNumber(Me.txtNumDays.Value) / 7, 2))
    UpdateFinalDateFromDays
    isRefreshing = False

    RefreshTotals
End Sub

Private Sub txtNumPeople_Change()
    RefreshTotals
End Sub

Private Sub txtStandardHours_Change()
    RefreshTotals
End Sub

Private Sub txtStandardRate_Change()
    RefreshTotals
End Sub

Private Sub txtOvertimeHours_Change()
    RefreshTotals
End Sub

Private Sub txtOvertimeRate_Change()
    RefreshTotals
End Sub

Private Sub txtWeekendHours_Change()
    RefreshTotals
End Sub

Private Sub txtPerDiem_Change()
    RefreshTotals
End Sub

Private Sub txtHotelRate_Change()
    RefreshTotals
End Sub

Private Sub txtNumCar_Change()
    RefreshTotals
End Sub

Private Sub txtCarRentRate_Change()
    RefreshTotals
End Sub

Private Sub txtFuelRate_Change()
    RefreshTotals
End Sub

Private Sub txtMileage_Change()
    RefreshTotals
End Sub

Private Sub txtMileageRate_Change()
    RefreshTotals
End Sub

Private Sub txtNumFlights_Change()
    RefreshTotals
End Sub

Private Sub txtFlightRate_Change()
    RefreshTotals
End Sub

' =========================================================
' Main workflow
' =========================================================

Private Sub HandleCreateAction()
    Dim ws As Worksheet
    Dim timeMaterialService As clsTimeMaterialService

    Dim data As Collection
    Dim errorMessage As String
    Dim success As Boolean

    Set ws = modUtils.GetTargetWorksheet()
    Set data = BuildTimeMaterialData()
    Set timeMaterialService = New clsTimeMaterialService

    SaveInputsToDrivers

    success = timeMaterialService.InsertTimeMaterialSection( _
        ws, _
        data, _
        errorMessage)

    If success Then
        MsgBox "T&M section inserted successfully.", vbInformation, FORM_TITLE
        Unload Me
    Else
        MsgBox errorMessage, vbCritical, FORM_TITLE
    End If
End Sub

' =========================================================
' UI setup
' =========================================================

Private Sub ConfigureForm()
    Me.Height = FORM_HEIGHT
    Me.Width = FORM_WIDTH
End Sub

' =========================================================
' Data collection
' =========================================================

Private Function BuildTimeMaterialData() As Collection
    Dim data As Collection
    Set data = New Collection

    data.Add Val(Me.txtNumPeople.Value), "numPeople"
    data.Add Val(Me.txtNumWeeks.Value), "numWeeks"
    data.Add Val(Me.txtNumDays.Value), "numDays"

    data.Add Val(Me.txtStandardHours.Value), "standardHours"
    data.Add Val(Me.txtStandardRate.Value), "standardRate"

    data.Add Val(Me.txtOvertimeHours.Value), "overtimeHours"
    data.Add Val(Me.txtOvertimeRate.Value), "overtimeRate"

    data.Add Val(Me.txtWeekendHours.Value), "weekendHours"

    data.Add Val(Me.txtPerDiem.Value), "perDiem"
    data.Add Val(Me.txtHotelRate.Value), "hotelRate"

    data.Add Val(Me.txtNumCar.Value), "numCar"
    data.Add Val(Me.txtCarRentRate.Value), "carRentRate"

    data.Add Val(Me.txtFuelRate.Value), "fuelRate"

    data.Add Val(Me.txtMileage.Value), "mileage"
    data.Add Val(Me.txtMileageRate.Value), "mileageRate"

    data.Add Val(Me.txtNumFlights.Value), "numFlights"
    data.Add Val(Me.txtFlightRate.Value), "flightRate"

    Set BuildTimeMaterialData = data
End Function

' =========================================================
' Calculation helpers
' =========================================================

Private Sub UpdateDurationFromDates()
    Dim dayCount As Double

    If Trim$(Me.txtStartDate.Value) = "" Then Exit Sub
    If Trim$(Me.txtFinalDate.Value) = "" Then Exit Sub

    dayCount = DateDiff("d", CDate(Me.txtStartDate.Value), CDate(Me.txtFinalDate.Value))

    isRefreshing = True
    Me.txtNumDays.Value = CStr(dayCount)
    Me.txtNumWeeks.Value = CStr(Round(dayCount / 7, 2))
    isRefreshing = False

    RefreshTotals
End Sub

Private Sub UpdateFinalDateFromDays()
    If Trim$(Me.txtStartDate.Value) = "" Then Exit Sub

    Me.txtFinalDate.Value = Format$(DateAdd("d", ToNumber(Me.txtNumDays.Value), CDate(Me.txtStartDate.Value)), "mm/dd/yyyy")
End Sub

Private Sub RefreshTotals()
    Dim numPeople As Double
    Dim numWeeks As Double
    Dim numDays As Double

    Dim standardHours As Double
    Dim standardRate As Double
    Dim overtimeHours As Double
    Dim overtimeRate As Double
    Dim weekendHours As Double

    Dim perDiem As Double
    Dim hotelRate As Double
    Dim numCar As Double
    Dim carRentRate As Double
    Dim fuelRate As Double
    Dim mileagePerDay As Double
    Dim mileageRate As Double
    Dim numFlights As Double
    Dim flightRate As Double

    Dim standardHoursTotal As Double
    Dim overtimeHoursTotal As Double
    Dim weekendHoursTotal As Double

    numPeople = ToNumber(Me.txtNumPeople.Value)
    numWeeks = ToNumber(Me.txtNumWeeks.Value)
    numDays = ToNumber(Me.txtNumDays.Value)

    standardHours = ToNumber(Me.txtStandardHours.Value)
    standardRate = ToNumber(Me.txtStandardRate.Value)

    overtimeHours = ToNumber(Me.txtOvertimeHours.Value)
    overtimeRate = ToNumber(Me.txtOvertimeRate.Value)

    weekendHours = ToNumber(Me.txtWeekendHours.Value)

    perDiem = ToNumber(Me.txtPerDiem.Value)
    hotelRate = ToNumber(Me.txtHotelRate.Value)

    numCar = ToNumber(Me.txtNumCar.Value)
    carRentRate = ToNumber(Me.txtCarRentRate.Value)

    fuelRate = ToNumber(Me.txtFuelRate.Value)

    mileagePerDay = ToNumber(Me.txtMileage.Value)
    mileageRate = ToNumber(Me.txtMileageRate.Value)

    numFlights = ToNumber(Me.txtNumFlights.Value)
    flightRate = ToNumber(Me.txtFlightRate.Value)

    standardHoursTotal = numWeeks * 5 * numPeople * standardHours
    overtimeHoursTotal = numWeeks * 5 * numPeople * overtimeHours
    weekendHoursTotal = numWeeks * numPeople * weekendHours

    Me.txtStandardHoursTotal.Value = standardHoursTotal
    Me.txtStandardCostTotal.Value = standardHoursTotal * standardRate

    Me.txtOvertimeHoursTotal.Value = overtimeHoursTotal
    Me.txtOvertimeCostTotal.Value = overtimeHoursTotal * overtimeRate

    Me.txtWeekendHoursTotal.Value = weekendHoursTotal
    Me.txtWeekendCostTotal.Value = weekendHoursTotal * overtimeRate

    Me.txtPerDiemTotal.Value = perDiem * numWeeks * numPeople
    Me.txtHotelTotal.Value = hotelRate * numDays * numPeople
    Me.txtCarRentTotal.Value = numCar * carRentRate * numWeeks
    Me.txtFuelTotal.Value = fuelRate * numWeeks
    Me.txtMileageTotal.Value = mileagePerDay * numPeople * numDays * mileageRate
    Me.txtFlightTotal.Value = flightRate * numFlights * 2
End Sub

Private Function ToNumber(ByVal value As Variant) As Double
    If Trim$(CStr(value)) = "" Then
        ToNumber = 0
    Else
        ToNumber = Val(value)
    End If
End Function


Private Sub LoadDefaultsFromDrivers()
    Dim wsDrivers As Worksheet

    Set wsDrivers = ThisWorkbook.Worksheets("Drivers")

    Me.txtNumPeople.Value = wsDrivers.Range("B13").Value
    Me.txtNumWeeks.Value = wsDrivers.Range("B14").Value
    Me.txtNumDays.Value = wsDrivers.Range("B15").Value

    Me.txtStandardHours.Value = wsDrivers.Range("B16").Value
    Me.txtOvertimeHours.Value = wsDrivers.Range("B17").Value
    Me.txtWeekendHours.Value = wsDrivers.Range("B18").Value

    Me.txtHotelRate.Value = wsDrivers.Range("B19").Value

    Me.txtStandardRate.Value = wsDrivers.Range("B20").Value
    Me.txtOvertimeRate.Value = wsDrivers.Range("B21").Value

    Me.txtNumFlights.Value = wsDrivers.Range("B22").Value
    Me.txtFlightRate.Value = wsDrivers.Range("B23").Value

    Me.txtPerDiem.Value = wsDrivers.Range("B24").Value

    Me.txtNumCar.Value = wsDrivers.Range("B25").Value
    Me.txtCarRentRate.Value = wsDrivers.Range("B26").Value

    Me.txtFuelRate.Value = wsDrivers.Range("B27").Value

    Me.txtMileage.Value = wsDrivers.Range("B28").Value
    Me.txtMileageRate.Value = wsDrivers.Range("B29").Value
End Sub

Private Sub SaveInputsToDrivers()
    Dim wsDrivers As Worksheet

    Set wsDrivers = ThisWorkbook.Worksheets("Drivers")

    wsDrivers.Range("B13").Value = Val(Me.txtNumPeople.Value)
    wsDrivers.Range("B14").Value = Val(Me.txtNumWeeks.Value)
    wsDrivers.Range("B15").Value = Val(Me.txtNumDays.Value)

    wsDrivers.Range("B16").Value = Val(Me.txtStandardHours.Value)
    wsDrivers.Range("B17").Value = Val(Me.txtOvertimeHours.Value)
    wsDrivers.Range("B18").Value = Val(Me.txtWeekendHours.Value)

    wsDrivers.Range("B19").Value = Val(Me.txtHotelRate.Value)

    wsDrivers.Range("B20").Value = Val(Me.txtStandardRate.Value)
    wsDrivers.Range("B21").Value = Val(Me.txtOvertimeRate.Value)

    wsDrivers.Range("B22").Value = Val(Me.txtNumFlights.Value)
    wsDrivers.Range("B23").Value = Val(Me.txtFlightRate.Value)

    wsDrivers.Range("B24").Value = Val(Me.txtPerDiem.Value)

    wsDrivers.Range("B25").Value = Val(Me.txtNumCar.Value)
    wsDrivers.Range("B26").Value = Val(Me.txtCarRentRate.Value)

    wsDrivers.Range("B27").Value = Val(Me.txtFuelRate.Value)

    wsDrivers.Range("B28").Value = Val(Me.txtMileage.Value)
    wsDrivers.Range("B29").Value = Val(Me.txtMileageRate.Value)
End Sub