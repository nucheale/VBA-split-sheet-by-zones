Private Function HEXconverter(hexValue As String) As String
    leftV = Left(hexValue, 2)
    rightV = Right(hexValue, 2)
    centerV = Right(Left(hexValue, 4), 2)
    HEXconverter = "&H" & rightV & centerV & leftV
End Function

Function twoDimArrayToOneDim(oldArr)
    Dim newArr As Variant
    ReDim newArr(1 To UBound(oldArr, 1) * UBound(oldArr, 2))
    For i = LBound(oldArr, 1) To UBound(oldArr, 1)
        newArr(i) = oldArr(i, 1)
    Next i
    twoDimArrayToOneDim = newArr
End Function

Sub zoneAdd()

    With Application
        .Calculation = xlCalculationManual
        .AskToUpdateLinks = False
        .DisplayAlerts = False
    End With

    Dim result#, sumWeightObject#, sumWeightZone#, sumWeight0#, sumWeight1#, km#, km0#, km1#, kmWeight As Double
    Dim weightsArr0() As Double, weightsArr1() As Double, weightsArrObjects() As Double
    Dim kmArr0() As Double, kmArr1() As Double
    Dim checkingSumWeight() As Double, checkingObjectsWeight() As Double, checkingLandfillsWeight() As Double 'массивы для проверки сумм масс (масса образования, масса 1 плеча, масса по полигонам. Масса 2 плеча и прямого вывоза не проверяется, т.к. если не сойдется масса по полигонам, то уже где-то ошибка)

    splitBlocksRows = 6

    ThisWorkbook.Sheets(1).Select
    Set mainWs = Sheets(1)
    With mainWs
        If .AutoFilterMode Then .AutoFilter.ShowAllData
        Set findcell0Start = .Range(.Cells(1, 1), .Cells(1000, 11)).Find("Прямой вывоз и первое плечо") '
        Set findcell0End = .Range(.Cells(findcell0Start.Row, 1), .Cells(1000, 1)).Find("Итого (масса образования)") '
        Set findCell1Start = .Range(.Cells(findcell0End.Row + 1, 1), .Cells(1000, 1)).Find("Первое плечо и прямой вывоз итоги") 'первое плечо и прямой вывоз итоги
        Set findCell1End = .Range(.Cells(findcell0End.Row + 1, 1), .Cells(1000, 1)).Find("Итого") 'первое плечо и прямой вывоз итоги
        Set findCell2Start = .Range(.Cells(findCell1End.Row + 1, 1), .Cells(1000, 1)).Find("Второе плечо")
        Set findCell2End = .Range(.Cells(findCell2Start.Row + 1, 1), .Cells(1000, 1)).Find("Итого")
        Set findCell3Start = .Range(.Cells(findCell2End.Row + 1, 1), .Cells(1000, 1)).Find("Объекты размещения") 'полигоны
        Set findCell3End = .Range(.Cells(findCell3Start.Row + 1, 1), .Cells(1000, 1)).Find("Итого") 'полигоны
    End With

    With Sheets("Справочник").ListObjects("Объекты")
        Dim coeffSort() As Double
        ReDim Preserve coeffSort(1 To .ListRows.Count)
        Dim objects As New Dictionary 'заполнение словаря из листа справочник + коэф сортировки
        For i = 2 To .ListRows.Count + 1
            If IsNumeric(Sheets(1).Cells(findCell1Start.Row + i, 7).Value) Then
                coeffSort(i - 1) = CDbl(Sheets(1).Cells(findCell1Start.Row + i, 7).Value) / CDbl(Sheets(1).Cells(findCell1Start.Row + i, 5).Value)
            End If
                objects.Add .Range.Cells(i, 1).Value, Array(.Range.Cells(i, 2), .Range.Cells(i, 3), .Range.Cells(i, 4), .Range.Cells(i, 5), .Range.Cells(i, 6), coeffSort(i - 1))
        Next i
    End With
    ' For Each x In objects.Keys
    '     Debug.Print x, "///", objects(x)(4)
    '     Debug.Print objects.Keys(1)
    ' Next x
    'Debug.Print objects.Item(objects.Keys(1))(0)
    'Debug.Print objects("Волхонка АО " & Chr(34) & "Невский экологический оператор" & Chr(34))(1)
    'Debug.Print objects(objects.Keys(1))(0)
    'Debug.Print objects(Sheets(1).Cells(341, 3).Value)(5)

    Dim zonesCells() As Variant
    Dim zonesAll() As Integer
    zonesCells = Range(Cells(findcell0Start.Row + 2, 1), Cells(findcell0End.Row - 1, 1))
    zone = 0
    For e = LBound(zonesCells) + 1 To UBound(zonesCells) 'находим количество лотов
        If Not zonesCells(e - 1, 1) = zonesCells(e, 1) Then
            ReDim Preserve zonesAll(1 To zone + 1)
            zonesAll(zone + 1) = CInt(zonesCells(e, 1))
            zone = zone + 1
        End If
    Next e
    
    zone = 1
    Dim zones() As Integer 'находим количество лотов удаление дубликатов
    ReDim Preserve zones(1 To 1)
    For Each e In zonesAll
        Unique = True
        For elem = LBound(zones) To UBound(zones)
            If zones(elem) = e Then
                Unique = False
                Exit For
            End If
        Next elem
        If Unique Then
            ReDim Preserve zones(1 To zone)
            zones(zone) = e
            zone = zone + 1
        End If
    Next e

    For Each zone In zones 'удаление листов если уже есть
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = "Лот " & zone Then ws.Delete
        Next ws
    Next zone

    For zone = LBound(zones) To UBound(zones)
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = "Лот " & zone
        Sheets(1).Select
        
        With Sheets("Лот " & zone)
            For i = 1 To findcell0Start.Row + 1 ''заголовки блок МО
                For j = 1 To 11
                    .Cells(i, j) = Sheets(1).Cells(i, j)
                Next j
            Next i

            startRow0 = findcell0Start.Row + 2 'начальная строка 1 блока без заголовков
            endRow0 = startRow0
            For i = 1 To findcell0End.Row 'первый блок с МО
                If i <= findcell0Start.Row + 1 Then
                    For j = 1 To 11
                        .Cells(i, j) = Sheets(1).Cells(i, j)
                    Next j
                ElseIf Sheets(1).Cells(i, 1) = zone Then
                    For j = 1 To 11
                        .Cells(endRow0, j) = Sheets(1).Cells(i, j)
                    Next j
                    endRow0 = endRow0 + 1
                End If
            Next i
            
            endRow0 = endRow0 - 1 'конечная строка 1 блока без итоговой

            Dim objectsArr As Variant, weightsArr As Variant, kmArr As Variant, objectsArrFull As Variant, weightsArrFull As Variant, kmArrFull As Variant

            objectsArr = .Range(.Cells(startRow0, 4), .Cells(endRow0, 4))
            weightsArr = .Range(.Cells(startRow0, 9), .Cells(endRow0, 9))
            kmArr = .Range(.Cells(startRow0, 10), .Cells(endRow0, 10))
            
            objectsArrFull = Sheets(1).Range(Cells(findcell0Start.Row + 2, 4), Cells(findcell0End.Row - 1, 4))
            weightsArrFull = Sheets(1).Range(Cells(findcell0Start.Row + 2, 9), Cells(findcell0End.Row - 1, 9))
            kmArrFull = Sheets(1).Range(Cells(findcell0Start.Row + 2, 10), Cells(findcell0End.Row - 1, 10))

            sumWeightZone = 0
            For e = LBound(weightsArr) To UBound(weightsArr)
                sumWeightZone = sumWeightZone + weightsArr(e, 1) 'масса образования по лоту
            Next e

            .Cells(endRow0 + 1, 1) = "Итого"
            .Cells(endRow0 + 1, 9) = sumWeightZone
            ReDim Preserve checkingSumWeight(1 To zone)
            checkingSumWeight(zone) = sumWeightZone

            sumMultiplicationResult = 0
            For e = 1 To UBound(weightsArr) 'средневзвешенное расстояние
                multiplicationResult = weightsArr(e, 1) * kmArr(e, 1)
                sumMultiplicationResult = sumMultiplicationResult + multiplicationResult
            Next e
            kmResult = sumMultiplicationResult / sumWeightZone
            .Cells(endRow0 + 1, 10) = kmResult 'средневзвешенное расстояние
            .Cells(endRow0 + 1, 11) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow0, 11), .Cells(endRow0, 11))) 'сумма РСО
            

            Dim sortPlaces() As String 'Наименования объектов прямого вывоза и 1 плеча
            ReDim sortPlaces(1 To 1)
            sortPlaces(1) = objectsArr(1, 1)
            countObj = 2
            For i = LBound(objectsArr, 1) + 1 To UBound(objectsArr, 1)
                If Not objectsArr(i - 1, 1) = objectsArr(i, 1) Then
                    ReDim Preserve sortPlaces(1 To countObj)
                    sortPlaces(countObj) = objectsArr(i, 1)
                    countObj = countObj + 1
                End If
            Next i
            
            startRow1 = endRow0 + splitBlocksRows 'начальная строка 2 блока (1 плечо и прямой вывоз итоги) без заголовков
            endrow1 = startRow1

            For j = 1 To 13 'заголовки 1 плечо
                .Cells(startRow1 - 2, j) = Sheets(1).Cells(findCell1Start.Row, findCell1Start.Column + j - 1)
                .Cells(startRow1 - 1, j) = Sheets(1).Cells(findCell1Start.Row + 1, findCell1Start.Column + j - 1)
            Next j

            sumWeight0 = 0
            sumWeight1 = 0
            counter0 = 0
            is0 = False
            For i = LBound(sortPlaces) To UBound(sortPlaces)
                sumWeightObject = 0
                resultKm = 0
                For e = LBound(objectsArr, 1) To UBound(objectsArr, 1) 'вес по объекту и суммпроизв расстояния
                    If objectsArr(e, 1) = sortPlaces(i) Then
                        sumWeightObject = sumWeightObject + weightsArr(e, 1)
                        resultKm = resultKm + (CDbl(weightsArr(e, 1)) * CDbl(kmArr(e, 1)))
                    End If
                Next e
                km = resultKm / sumWeightObject

                .Cells(endrow1, 1) = objects(sortPlaces(i))(4) 'плечо объекта
                .Cells(endrow1, 3) = sortPlaces(i) 'название
                .Cells(endrow1, 5) = sumWeightObject 'масса на объект
                .Cells(endrow1, 10) = km 'сревзв расстояние
                
                If objects(sortPlaces(i))(4) = "Первое плечо" Then
                    .Cells(endrow1, 6) = objects(sortPlaces(i))(0) 'отбор ВМР
                    sortLimit = sumWeightObject * objects(sortPlaces(i))(5) 'производственная программа
                    .Cells(endrow1, 7) = sortLimit
                    .Cells(endrow1, 8) = sumWeightObject - Application.WorksheetFunction.Min(sumWeightObject, sortLimit * objects(sortPlaces(i))(0)) 'масса после обраб
                    .Cells(endrow1, 9) = Application.WorksheetFunction.Min(sumWeightObject, sortLimit * objects(sortPlaces(i))(0))
                    If sumWeightObject > sortLimit Then .Cells(endrow1, 13) = Round(sumWeightObject - sortLimit, 6) Else: .Cells(endrow1, 13) = 0 'перегруз
                    sumWeight1 = sumWeight1 + sumWeightObject 'общая масса 1 плеча
                ElseIf objects(sortPlaces(i))(4) = "Прямой вывоз" Then
                    is0 = True
                    .Cells(endrow1, 6) = "—"
                    .Cells(endrow1, 7) = "—"
                    .Cells(endrow1, 8) = "—"
                    .Cells(endrow1, 9) = "—"
                    .Cells(endrow1, 13) = 0 'перегруз
                    sumWeight0 = sumWeight0 + sumWeightObject 'общая масса прямой вывоз
                    counter0 = counter0 + 1
                End If
                endrow1 = endrow1 + 1
            Next i
            endrow1 = endrow1 - 1
            
            .Cells(endrow1, 12) = sumWeight1 'масса 1 плечо
            If is0 Then .Cells(startRow1 + counter0 - 1, 12) = sumWeight0 'масса прямого вывоза

            ReDim Preserve checkingObjectsWeight(1 To zone)
            checkingObjectsWeight(zone) = sumWeight1 + sumWeight0 'масса 1 плеча и прямого вывоза

            typeArr01 = .Range(.Cells(startRow1, 1), .Cells(endrow1, 1))
            typeArr01 = twoDimArrayToOneDim(typeArr01)
            weightsArr01 = .Range(.Cells(startRow1, 5), .Cells(endrow1, 5))
            weightsArr01 = twoDimArrayToOneDim(weightsArr01)
            kmArr01 = .Range(.Cells(startRow1, 10), .Cells(endrow1, 10))
            kmArr01 = twoDimArrayToOneDim(kmArr01)
            
            km1 = 0
            For e = LBound(typeArr01) To UBound(typeArr01)
                If typeArr01(e) = "Первое плечо" Then
                    kmWeight = weightsArr01(e) * kmArr01(e)
                    km1 = km1 + kmWeight
                End If
            Next e
            km1 = km1 / sumWeight1
            .Cells(endrow1, 11) = km1 'средневзвешенное 1 плечо

            If is0 Then
                km0 = 0
                For e = LBound(typeArr01) To UBound(typeArr01)
                    If typeArr01(e) = "Прямой вывоз" Then
                        kmWeight = weightsArr01(e) * kmArr01(e)
                        km0 = km0 + kmWeight
                    End If
                Next e
                km0 = km0 / sumWeight0
                .Cells(startRow1 + counter0 - 1, 11) = km0 'средневзвешенное прямой вывоз
            End If

            .Cells(endrow1 + 1, 1) = "Итого"
            .Cells(endrow1 + 1, 5) = sumWeightZone
            .Cells(endrow1 + 1, 7) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow1, 7), .Cells(endrow1, 7)))
            .Cells(endrow1 + 1, 8) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow1, 8), .Cells(endrow1, 8)))
            .Cells(endrow1 + 1, 9) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow1, 9), .Cells(endrow1, 9)))
            .Cells(endrow1 + 1, 13) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow1, 13), .Cells(endrow1, 13)))


            Dim weightsAfterSort() As Double
            Dim objectsAfterSort() As String
            For i = startRow1 To endrow1
                ReDim Preserve weightsAfterSort(1 To i - startRow1 + 1) 'масса после обработки 1 плечо
                ReDim Preserve objectsAfterSort(1 To i - startRow1 + 1) 'наименования объектов 1 плечо
                If IsNumeric(.Cells(i, 8)) = True Then weightsAfterSort(i - startRow1 + 1) = .Cells(i, 8) Else weightsAfterSort(i - startRow1 + 1) = 0
                objectsAfterSort(i - startRow1 + 1) = .Cells(i, 3)
            Next i

            km = 0
            For i = startRow1 To endrow1 'средневзвешенное итог
                kmWeight = weightsArr01(i - startRow1 + 1) * kmArr01(i - startRow1 + 1)
                km = km + kmWeight
            Next i
            km = km / sumWeightZone
            .Cells(endrow1 + 1, 10) = km

            .Range(.Cells(startRow1, 1), .Cells(endrow1, 13)).Sort Key1:=.Range(.Cells(startRow1, 1), .Cells(endrow1, 1)), Order1:=xlAscending, Header:=xlNo 'сортировка чтобы сначала 1 плечо, потом прямой вывоз

            For i = startRow1 To endrow1 'объединение ячеек расстояний и масс
                If Not .Cells(i, 11) = Empty Then
                    .Range(.Cells(startRow1, 11), .Cells(i, 11)).MergeCells = 1
                    .Range(.Cells(startRow1, 12), .Cells(i, 12)).MergeCells = 1
                    mergedRow = i
                    For ii = mergedRow + 1 To endrow1
                        If .Cells(ii, 11) = Empty Then
                            .Range(.Cells(ii, 11), .Cells(endrow1, 11)).MergeCells = 1
                            .Range(.Cells(ii, 12), .Cells(endrow1, 12)).MergeCells = 1
                            Exit For
                        End If
                    Next ii
                    Exit For
                End If
            Next i
            '--------------------------------Конец 1 плеча--------------------------------------
        
            Dim sumWeightObjectsZone() As Variant, sumWeightObjectsFull() As Double
            ReDim Preserve sumWeightObjectsZone(1 To objects.Count) 'суммы масс на объекты по лоту (если объекта нет, ставится 0) (16 шт)
            ReDim Preserve sumWeightObjectsFull(1 To objects.Count) 'суммы масс на объекты по всем лотам (16 шт)
            
            For e = 0 To objects.Count - 1
                For i = startRow1 To endrow1
                    If objects.Keys(e) = .Cells(i, 3) Then sumWeightObjectsZone(e + 1) = .Cells(i, 5)
                Next i

                If sumWeightObjectsZone(e + 1) = Empty Then sumWeightObjectsZone(e + 1) = 0

                For element = LBound(objectsArrFull) To UBound(objectsArrFull)
                    If objectsArrFull(element, 1) = objects.Keys(e) Then sumWeightObjectsFull(e + 1) = sumWeightObjectsFull(e + 1) + weightsArrFull(element, 1)
                Next element
            Next e
        
            '  -----------------------------------2 плечо-----------------------------------------

            startRow2 = endrow1 + splitBlocksRows 'начальная строка 3 блока (2 плечо) без заголовков
            
            For j = 1 To 12 'заголовки
                .Cells(startRow2 - 2, j) = Sheets(1).Cells(findCell2Start.Row, findCell2Start.Column + j - 1)
                .Cells(startRow2 - 1, j) = Sheets(1).Cells(findCell2Start.Row + 1, findCell2Start.Column + j - 1)
            Next j
            startRow = findCell2Start.Row + 1

            ' Set findCellSheet1 = Sheets(1).Range(Cells(startRow, 1), Cells(1000, 1)).Find("Итого")
            endrow = findCell2End.Row - 1
            endRow2 = startRow2 + (endrow - startRow - 1)
            
            For i = 1 To (endrow - startRow)
                .Cells(startRow2 + i - 1, 1) = "Второе плечо"
                .Cells(startRow2 + i - 1, 3) = Sheets(1).Cells(startRow + i, 3)
                .Cells(startRow2 + i - 1, 6) = Sheets(1).Cells(startRow + i, 6)
                .Cells(startRow2 + i - 1, 8) = Sheets(1).Cells(startRow + i, 8)
            Next i

            counter = 0
            For i = startRow + 1 To endrow 'масса на объект по лоту в списке всего 2 плеча
                For e = 0 To objects.Count - 1
                    If Sheets(1).Cells(i, 6) = objects.Keys(e) Then
                        .Cells(startRow2 + counter, 5) = Sheets(1).Cells(i, 5) * (sumWeightObjectsZone(e + 1) / sumWeightObjectsFull(e + 1))
                        .Cells(startRow2 + counter, 9) = Sheets(1).Cells(i, 9) * (sumWeightObjectsZone(e + 1) / sumWeightObjectsFull(e + 1))
                        .Cells(startRow2 + counter, 10) = Sheets(1).Cells(i, 10) * (sumWeightObjectsZone(e + 1) / sumWeightObjectsFull(e + 1))
                        counter = counter + 1
                    End If
                Next e
            Next i

            For i = endRow2 To (endRow2 - (endrow - startRow)) Step -1
                If .Cells(i, 5) = 0 Then .Rows(i).EntireRow.Delete 'удаление объектов с массой 0
            Next i
            endRow2 = .Cells(Rows.Count, 1).End(xlUp).Row 'конечная строка 3 блока (2 плечо) без итоговой
            
            Dim landfills2(), km2(), weights2(), weights2ByZone(), sortedWeights2(), sortedWeights2ByZone(), unsortedWeights2(), unsortedWeights2ByZone()
            ReDim landfills2(1 To (endRow2 - startRow2 + 1))
            ReDim km2(1 To (endRow2 - startRow2 + 1))
            ReDim weights2(1 To (endRow2 - startRow2 + 1))
            ReDim sortedWeights2(1 To (endRow2 - startRow2 + 1))
            ReDim unsortedWeights2(1 To (endRow2 - startRow2 + 1))

            weights2Sum = 0
            sortedWeights2Sum = 0
            unsortedWeights2Sum = 0
            For i = startRow2 To endRow2 'полигоны, расстояния, массы 2 плеча, сорта/несорта + их сумма
                landfills2(i - startRow2 + 1) = .Cells(i, 3)
                km2(i - startRow2 + 1) = .Cells(i, 8)
                weights2(i - startRow2 + 1) = .Cells(i, 5)
                weights2Sum = weights2Sum + weights2(i - startRow2 + 1)
                sortedWeights2(i - startRow2 + 1) = .Cells(i, 9)
                sortedWeights2Sum = sortedWeights2Sum + sortedWeights2(i - startRow2 + 1)
                unsortedWeights2(i - startRow2 + 1) = .Cells(i, 10)
                unsortedWeights2Sum = unsortedWeights2Sum + unsortedWeights2(i - startRow2 + 1)
            Next i
            
            iterationN = 1
            mergedRow = startRow2
            For i = startRow2 To endRow2
                For e = LBound(weightsAfterSort) To UBound(weightsAfterSort)
                    If .Cells(i, 6) = objectsAfterSort(e) Then .Cells(i, 11) = weights2(endRow2 - startRow2 + 1) / CDbl(weightsAfterSort(e)) 'доля от общего потока объекта
                Next e

                sumMultiplicationResult = 0
                If Not .Cells(i, 3) = .Cells(i + 1, 3) Then 'средневзвешенные расстояния
                    weightLandfill = 0
                    For e = LBound(weights2) To UBound(weights2)
                        If landfills2(e) = .Cells(i, 3) Then
                            weightLandfill = weightLandfill + weights2(e)
                            multiplicationResult = weights2(e) * km2(e)
                            sumMultiplicationResult = sumMultiplicationResult + multiplicationResult
                        End If
                    Next e
                    km2Result = sumMultiplicationResult / weightLandfill
                    .Cells(i, 12) = km2Result
                End If

                If Not .Cells(i, 12) = Empty Then 'объединение ячеек средневзвешенных расстояний
                    If iterationN = 1 Then
                        iterationN = iterationN + 1
                        GoTo continueFor
                    End If
                    .Range(.Cells(mergedRow, 12), .Cells(i, 12)).MergeCells = 1
                    mergedRow = i + 1
                End If
                iterationN = iterationN + 1
continueFor:
            Next i

            ReDim Preserve weights2ByZone(1 To zone) 'для проверки суммы в конце (сумма массы 2 плеча, сумма сорта и несорта)
            ReDim Preserve sortedWeights2ByZone(1 To zone)
            ReDim Preserve unsortedWeights2ByZone(1 To zone)
            weights2ByZone(zone) = weights2Sum
            sortedWeights2ByZone(zone) = sortedWeights2Sum
            unsortedWeights2ByZone(zone) = unsortedWeights2Sum
            
            .Cells(endRow2 + 1, 1) = "Итого"
            .Cells(endRow2 + 1, 5) = weights2Sum
            .Cells(endRow2 + 1, 9) = sortedWeights2Sum
            .Cells(endRow2 + 1, 10) = unsortedWeights2Sum

            Dim km2SumMul#
            km2SumMul = 0 'средневзвешенное итоговое по 2 плечу
            For i = LBound(weights2) To UBound(weights2)
                km2SumMul = km2SumMul + (weights2(i) * km2(i))
            Next i
            .Cells(endRow2 + 1, 12) = km2SumMul / weights2Sum


            '-----------------------------------Конец 2 плечо-----------------------------------------


            '-----------------------------------Средневзвешенные по полигонам-----------------------------------------

            '-----------------------------------Конец средневзвешенные по полигонам-----------------------------------------
            

            '-----------------------------------Полигоны-----------------------------------------

            Dim landfillsList() As String 'список полигоное без дубликатов из словаря

            element = 1
            For Key = 0 To objects.Count - 1 'список полигоное без дубликатов из словаря
            'Debug.Print objects(objects.Keys(key))(4)
                If objects(objects.Keys(Key))(4) = "Прямой вывоз" Then
                    ReDim Preserve landfillsList(1 To element)
                    landfillsList(element) = objects.Keys(Key)
                    element = element + 1
                End If
            Next Key


            Set findCell = Sheets(1).Range(Cells(1, 1), Cells(1000, 1)).Find("Объекты размещения")

            startRow3 = endRow2 + splitBlocksRows 'начальная строка 3 блока (Объекты размещения) без заголовков
            endRow3 = startRow3 + UBound(landfillsList) - 1 'конечная строка 3 блока (Объекты размещения) без итогов
            .Cells(startRow3 - 2, 1) = Sheets(1).Cells(findCell.Row, 1)

            For j = 1 To 11
                .Cells(startRow3 - 1, j) = Sheets(1).Cells(findCell.Row + 1, j) 'заголовки
            Next j

            For e = 1 To UBound(landfillsList) 'названия полигонов (заполняются все, даже если на них 0 т)
                .Cells(startRow3 + e - 1, 1) = landfillsList(e)
            Next e

            Dim sumWeights0Landfills() As Double, sumWeights2Landfills() As Double 'веса полигонов по прямому вывозу и 2 плечу
            ReDim Preserve sumWeights0Landfills(1 To UBound(landfillsList))
            ReDim Preserve sumWeights2Landfills(1 To UBound(landfillsList))

            For e = 1 To UBound(landfillsList)
                sumWeightLandfill0 = 0
                For i = startRow1 To endrow1 'цикл по блоку 1 плечо и прямой вывоз итоги
                    If landfillsList(e) = .Cells(i, 3) Then sumWeightLandfill0 = sumWeightLandfill0 + .Cells(i, 5) 'суммарный вес прямого вывоза по одному полигону
                Next i
                sumWeights0Landfills(e) = sumWeightLandfill0 'веса полигонов прямой вывоз
                
                sumWeightLandfill2 = 0
                For i = startRow2 To endRow2 'цикл по блоку 2 плечо
                    If landfillsList(e) = .Cells(i, 3) Then sumWeightLandfill2 = sumWeightLandfill2 + .Cells(i, 5) 'суммарный вес 2 плеча по одному полигону
                Next i
                sumWeights2Landfills(e) = sumWeightLandfill2 'веса полигонов 2 плечо
            Next e

            For e = 1 To UBound(landfillsList)
                .Cells(startRow3 + e - 1, 4) = sumWeights0Landfills(e)
                .Cells(startRow3 + e - 1, 5) = sumWeights2Landfills(e)
                .Cells(startRow3 + e - 1, 6) = sumWeights0Landfills(e) + sumWeights2Landfills(e)

                For Key = 0 To objects.Count - 1
                    If objects.Keys(Key) = landfillsList(e) Then
                        .Cells(startRow3 + e - 1, 7) = objects(objects.Keys(Key))(0) '% ВМР
                        .Cells(startRow3 + e - 1, 8) = objects(objects.Keys(Key))(1) 'лимит обработки
                        weightResult = (sumWeights0Landfills(e) + sumWeights2Landfills(e)) - Application.WorksheetFunction.Min((sumWeights0Landfills(e) + sumWeights2Landfills(e)), objects(objects.Keys(Key))(1)) * objects(objects.Keys(Key))(0) 'масса размещения
                        .Cells(startRow3 + e - 1, 9) = weightResult
                        .Cells(startRow3 + e - 1, 10) = objects(objects.Keys(Key))(2) 'лимит размещения
                        .Cells(startRow3 + e - 1, 11) = weightResult / objects(objects.Keys(Key))(2) 'загрузка объекта размещения
                    End If
                Next Key
            Next e

            Set findCellSheet1 = Sheets(1).Range(Cells(1, 1), Cells(1000, 1)).Find("Объект размещения")

            Dim landfillsWeightFull() As Double, landfillsWeightZone() As Double, coeffLandfills() As Double
            ReDim Preserve landfillsWeightZone(1 To UBound(landfillsList))
            ReDim Preserve landfillsWeightFull(1 To UBound(landfillsList))
            ReDim Preserve coeffLandfills(1 To UBound(landfillsList))

            For e = 1 To UBound(landfillsList) 'коэф. лота в полигоне
                For i = startRow3 To endRow3
                    If landfillsList(e) = .Cells(i, 1) Then
                        landfillsWeightZone(e) = CDbl(.Cells(i, 6)) 'общий вес полигона по текущему лоту (поступление)
                        Exit For
                    End If
                Next i
                For i = findCellSheet1.Row + 1 To findCellSheet1.Row + 1 + UBound(landfillsList)
                    If landfillsList(e) = Sheets(1).Cells(i, 1) Then
                        landfillsWeightFull(e) = CDbl(Sheets(1).Cells(i, 6)) 'общий вес полигона по всем лотам (поступление)
                        Exit For
                    End If
                Next i
                coeffLandfills(e) = landfillsWeightZone(e) / landfillsWeightFull(e) 'коэф. лота в полигоне. Возможно правильнее добавить в словарь
                ' Debug.Print coeffLandfills(e)
            Next e

            For e = 1 To UBound(landfillsWeightZone) 'еще раз заполнение лимита обаботки с учетом коэффициента лота в полигоне
                .Cells(startRow3 + e - 1, 8) = .Cells(startRow3 + e - 1, 8) * coeffLandfills(e)
            Next e

            .Cells(endRow3 + 1, 1) = "Итого" 'строка с итогами
            .Cells(endRow3 + 1, 4) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow3, 4), .Cells(endRow3, 4)))
            .Cells(endRow3 + 1, 5) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow3, 5), .Cells(endRow3, 5)))
            .Cells(endRow3 + 1, 6) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow3, 6), .Cells(endRow3, 6)))
            .Cells(endRow3 + 1, 8) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow3, 8), .Cells(endRow3, 8)))
            .Cells(endRow3 + 1, 9) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow3, 9), .Cells(endRow3, 9)))
            .Cells(endRow3 + 1, 10) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow3, 10), .Cells(endRow3, 10)))
            .Cells(endRow3 + 1, 11) = .Cells(endRow3 + 1, 9) / .Cells(endRow3 + 1, 10)

            ReDim Preserve checkingLandfillsWeight(1 To zone)
            checkingLandfillsWeight(zone) = .Cells(endRow3 + 1, 6)

            For i = endRow3 To startRow3 Step -1 'удаляем ненужные полигоны
                If .Cells(i, 6) = 0 Then
                    .Rows(i).EntireRow.Delete
                    endRow3 = endRow3 - 1
                End If
            Next i
            '-----------------------------------Конец полигоны-----------------------------------------

            ' For e = 0 To objects.Count - 1
            '     Debug.Print objects(objects.Keys(e))(5)
            ' Next e

            Erase sumWeightObjectsZone 'очищаем т.к. redim оставляет первое значение этих массивов и прибавляет к ним новые значения
            Erase sumWeightObjectsFull 'очищаем т.к. redim оставляет первое значение этих массивов и прибавляет к ним новые значения


            '-----------------------------------Форматирование-----------------------------------------
        
            .Cells.Font.Name = "Arial"
            .Cells.Font.Size = 11
            .Cells.HorizontalAlignment = xlCenter
            .Cells.VerticalAlignment = xlCenter

            For col = 1 To 15
                colWidth = Sheets(1).Columns(col).ColumnWidth
                .Columns(col).ColumnWidth = colWidth
            Next col

            .Rows(1).RowHeight = 20
            .Rows(2).RowHeight = 20
            .Rows(endRow0 + 1).RowHeight = 20
            .Range(.Cells(1, 1), .Cells(1, 11)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(1, 1), .Cells(1, 11)).Interior.Color = HEXconverter("DFEDEF")
            .Range(.Cells(2, 1), .Cells(2, 11)).Interior.Color = &HD9D9D9 '#D9D9D9
            '.Range(.Cells(4, 2), .Cells(endrow, 2)).Interior.Color = &HB1B1FF'#B1B1FF
            .Range(.Cells(2, 1), .Cells(2, 11)).Merge
            .Range(.Cells(1, 1), .Cells(3, 11)).Font.Bold = True
            .Range(.Cells(3, 4), .Cells(endRow0 + 1, 5)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(3, 7), .Cells(endRow0 + 1, 8)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(1, 1), .Cells(endRow0 + 1, 11)).borders.LineStyle = xlContinuous
            .Range(.Cells(4, 9), .Cells(endRow0 + 1, 11)).NumberFormat = "#,##0.00"
            .Range(.Cells(endRow0 + 1, 1), .Cells(endRow0 + 1, 11)).Font.Bold = True
            .Range(.Cells(endRow0 + 1, 1), .Cells(endRow0 + 1, 8)).HorizontalAlignment = xlCenterAcrossSelection
            
            .Rows(startRow1 - 2).RowHeight = 20
            .Range(.Cells(startRow1 - 2, 1), .Cells(startRow1 - 2, 13)).Interior.Color = &HD9D9D9 '#D9D9D9
            .Range(.Cells(startRow1 - 2, 1), .Cells(startRow1 - 2, 13)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow1 - 1, 1), .Cells(endrow1, 2)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow1 - 1, 3), .Cells(endrow1, 4)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow1 - 1, 10), .Cells(startRow1 - 1, 11)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow1 - 2, 1), .Cells(startRow1 - 1, 13)).Font.Bold = True
            .Range(.Cells(startRow1 - 2, 1), .Cells(endrow1 + 1, 13)).borders.LineStyle = xlContinuous
            .Range(.Cells(startRow1, 5), .Cells(endrow1 + 1, 5)).NumberFormat = "#,##0.00"
            .Range(.Cells(startRow1, 6), .Cells(endrow1 + 1, 6)).NumberFormat = "0.00%"
            .Range(.Cells(startRow1, 7), .Cells(endrow1 + 1, 13)).NumberFormat = "#,##0.00"
            .Range(.Cells(endrow1 + 1, 1), .Cells(endrow1 + 1, 13)).Font.Bold = True
            .Range(.Cells(endrow1 + 1, 1), .Cells(endrow1 + 1, 4)).HorizontalAlignment = xlCenterAcrossSelection

            .Rows(startRow2 - 2).RowHeight = 20
            .Range(.Cells(startRow2 - 2, 1), .Cells(startRow2 - 2, 12)).Interior.Color = &HD9D9D9 '#D9D9D9
            .Range(.Cells(startRow2 - 2, 1), .Cells(startRow2 - 2, 12)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow2 - 1, 1), .Cells(endRow2, 2)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow2 - 1, 3), .Cells(endRow2, 4)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow2 - 1, 6), .Cells(endRow2, 7)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow2 - 2, 1), .Cells(startRow2 - 1, 12)).Font.Bold = True
            .Range(.Cells(startRow2 - 2, 1), .Cells(endRow2 + 1, 12)).borders.LineStyle = xlContinuous
            .Range(.Cells(startRow2, 5), .Cells(endRow2 + 1, 5)).NumberFormat = "#,##0.00"
            .Range(.Cells(startRow2, 8), .Cells(endRow2 + 1, 10)).NumberFormat = "#,##0.00"
            .Range(.Cells(startRow2, 11), .Cells(endRow2 + 1, 11)).NumberFormat = "0.00%"
            .Range(.Cells(startRow2, 12), .Cells(endRow2 + 1, 12)).NumberFormat = "#,##0.00"
            .Range(.Cells(endRow2 + 1, 1), .Cells(endRow2 + 1, 12)).Font.Bold = True
            .Range(.Cells(endRow2 + 1, 1), .Cells(endRow2 + 1, 4)).HorizontalAlignment = xlCenterAcrossSelection

            '.rows(startRow3-2).rowHeight = 20
            .Range(.Cells(startRow3 - 2, 1), .Cells(startRow3 - 2, 11)).Interior.Color = &HD9D9D9 '#D9D9D9
            .Range(.Cells(startRow3 - 2, 1), .Cells(startRow3 - 2, 11)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow3 - 1, 1), .Cells(endRow3 + 1, 3)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow3 - 2, 1), .Cells(startRow3 - 1, 11)).Font.Bold = True
            .Range(.Cells(startRow3 - 2, 1), .Cells(endRow3 + 1, 11)).borders.LineStyle = xlContinuous
            .Range(.Cells(startRow3, 4), .Cells(endRow3 + 1, 6)).NumberFormat = "#,##0.00"
            .Range(.Cells(startRow3, 7), .Cells(endRow3 + 1, 13)).NumberFormat = "0.00%"
            .Range(.Cells(startRow3, 8), .Cells(endRow3 + 1, 10)).NumberFormat = "#,##0.00"
            .Range(.Cells(startRow3, 11), .Cells(endRow3 + 1, 11)).NumberFormat = "0.00%"
            .Range(.Cells(endRow3 + 1, 1), .Cells(endRow3 + 1, 11)).Font.Bold = True

            Dim h20(1 To 5) As Long 'высота строк
            h20(1) = endRow0 + 1
            h20(2) = startRow0 - 2
            h20(3) = startRow1 - 2
            h20(4) = startRow2 - 2
            h20(5) = startRow3 - 2

            For e = 1 To UBound(h20)
                .Rows(h20(e)).RowHeight = 28
            Next e
            For i = startRow1 To endrow1 + 1
                .Rows(i).RowHeight = 28
            Next i
            For i = startRow2 To endRow2 + 1
                .Rows(i).RowHeight = 28
            Next i
            For i = startRow3 To endRow3 + 1
                .Rows(i).RowHeight = 28
            Next i

            .Cells.WrapText = True
            
            Dim type0 As Variant
            type0 = .Range(.Cells(1, 2), .Cells(endRow0, 2))
            For e = LBound(type0) To UBound(type0)
                Select Case type0(e, 1)
                    Case "ОР"
                        .Cells(e, 2).Interior.Color = HEXconverter("B1B1FF")
                    Case "МСС"
                        .Cells(e, 2).Interior.Color = HEXconverter("E8FCAA")
                    Case "МПС"
                        .Cells(e, 2).Interior.Color = HEXconverter("A4B5B6")
                End Select
            Next e

            Dim sort1 As Variant
            sort1 = .Range(.Cells(startRow1, 6), .Cells(endrow1, 6))
            For e = LBound(sort1) To UBound(sort1)
                Select Case sort1(e, 1)
                    Case "—"
                        .Cells(startRow1 + e - 1, 6).Interior.Color = HEXconverter("B1B1FF")
                    Case 0
                        .Cells(startRow1 + e - 1, 6).Interior.Color = HEXconverter("A4B5B6")
                    Case Is > 0
                        .Cells(startRow1 + e - 1, 6).Interior.Color = HEXconverter("E8FCAA")
                End Select
            Next e
            
            'Debug.Print endRow0


        End With
        '-----------------------------------Конец Форматирование-----------------------------------------
        
        '-----------------------------------Еще немного форматирования-----------------------------------------
        ThisWorkbook.Worksheets("Лот " & zone).Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 75
        
        '-----------------------------------Конец Еще немного форматирования-----------------------------------------

    Next zone

    '-----------------------------------Проверка-----------------------------------------

    Dim checkBad As Boolean
    checkBad = False
    Dim errText$
    
    For zone = LBound(zones) To UBound(zones)
        resultCheckingSumWeight = resultCheckingSumWeight + checkingSumWeight(zone)
        resultCheckingObjectsWeight = resultCheckingObjectsWeight + checkingObjectsWeight(zone)
        resultCheckingLandfillsWeight = resultCheckingLandfillsWeight + checkingLandfillsWeight(zone)
        If Not WorksheetFunction.Round(weights2ByZone(zone), 4) = WorksheetFunction.Round((sortedWeights2ByZone(zone) + unsortedWeights2ByZone(zone)), 4) Then
            checkBad = True
            errText = errText & "Лот " & zone & ": " & vbLf & "Масса 2 плеча: " & WorksheetFunction.Round(weights2ByZone(zone), 4) _
            & vbLf & "Масса 2 плеча как сорт + несорт: " & WorksheetFunction.Round((sortedWeights2ByZone(zone) + unsortedWeights2ByZone(zone)), 4) & vbLf & vbLf
        End If
    Next zone

    If Not WorksheetFunction.Round(resultCheckingSumWeight, 5) = WorksheetFunction.Round(Sheets(1).Cells(findcell0End.Row, 9), 5) Then
        checkBad = True
        errText = errText & "Масса образования расчет: " & WorksheetFunction.Round(resultCheckingSumWeight, 5) & vbLf _
        & "Масса образования исходная: " & WorksheetFunction.Round(Sheets(1).Cells(findcell0End.Row, 9), 5) & vbLf & vbLf
    End If
    If Not WorksheetFunction.Round(resultCheckingObjectsWeight, 5) = WorksheetFunction.Round(Sheets(1).Cells(findCell1End.Row, 5), 5) Then
        checkBad = True
        errText = errText & "Масса 1 плеча и прямого вывоза расчет: " & WorksheetFunction.Round(resultCheckingObjectsWeight, 5) & vbLf _
        & "Масса 1 плеча и прямого вывоза исходная: " & WorksheetFunction.Round(Sheets(1).Cells(findCell1End.Row, 5), 5) & vbLf & vbLf
    End If
    If Not WorksheetFunction.Round(resultCheckingLandfillsWeight, 5) = WorksheetFunction.Round(Sheets(1).Cells(findCell3End.Row, 6), 5) Then
        checkBad = True
        errText = errText & "Масса полигонов расчет: " & WorksheetFunction.Round(resultCheckingLandfillsWeight, 5) & vbLf _
        & "Масса полигонов исходная: " & WorksheetFunction.Round(Sheets(1).Cells(findCell3End.Row, 6), 5)
    End If

    If checkBad = True Then endText = "Обнаружены ошибки. " & vbLf & vbLf & errText Else: endText = "Все проверки пройдены успешно"
    MsgBox endText

    '-----------------------------------Конец Проверка-----------------------------------------
    
    With Application
        .AskToUpdateLinks = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With

End Sub



