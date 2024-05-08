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
    Dim weightsArr01(), weightsArr0(), weightsArr1(), weightsArrObjects() As Double
    Dim kmArr01(), kmArr0(), kmArr1() As Double
    Dim typeArr01() As String
    Dim checkingSumWeight(), checkingObjectsWeight(), checkingLandfillsWeight() As Double 'массивы для проверки сумм масс (масса образования, масса 1 плеча, масса по полигонам. Масса 2 плеча и прямого вывоза не проверяется, т.к. если не сойдется масса по полигонам, то уже где-то ошибка)

    Sheets(1).Select

    Set findcell0Start = Sheets(1).Range(Cells(1, 1), Cells(1000, 11)).Find("Прямой вывоз и первое плечо") '
    Set findcell0End = Sheets(1).Range(Cells(findcell0Start.Row, 1), Cells(1000, 1)).Find("Итого (масса образования)") '
    Set findCell1Start = Sheets(1).Range(Cells(findcell0End.Row + 1, 1), Cells(1000, 1)).Find("Первое плечо и прямой вывоз итоги") 'первое плечо и прямой вывоз итоги
    Set findCell1End = Sheets(1).Range(Cells(findcell0End.Row + 1, 1), Cells(1000, 1)).Find("Итого") 'первое плечо и прямой вывоз итоги
    Set findCell2Start = Sheets(1).Range(Cells(findCell1End.Row + 1, 1), Cells(1000, 1)).Find("Второе плечо")
    Set findCell2End = Sheets(1).Range(Cells(findCell2Start.Row + 1, 1), Cells(1000, 1)).Find("Итого")
    Set findCell3Start = Sheets(1).Range(Cells(findCell2End.Row + 1, 1), Cells(1000, 1)).Find("Объекты размещения") 'полигоны
    Set findCell3End = Sheets(1).Range(Cells(findCell3Start.Row + 1, 1), Cells(1000, 1)).Find("Итого") 'полигоны


    Set findCell = Nothing
    Set findCell = Sheets(1).Range(Cells(1, 1), Cells(1000, 1)).Find("Первое плечо и прямой вывоз итоги")

    Dim coeffSort() As Double
                
    Dim objects As New Dictionary
    With objects
        For i = 2 To Sheets("Справочник").ListObjects("Объекты").ListRows.Count + 1
            ReDim Preserve coeffSort(1 To (i - 1))
            If IsNumeric(Sheets(1).Cells(findCell.Row + 1 + (i - 1), 7).Value) Then
                coeffSort(i - 1) = CDbl(Sheets(1).Cells(findCell.Row + 1 + (i - 1), 7).Value) / CDbl(Sheets(1).Cells(findCell.Row + 1 + (i - 1), 5).Value)
            End If
            'Debug.Print coeffSort((i - 1))
            
            objects.Add Sheets("Справочник").ListObjects("Объекты").Range.Cells(i, 1).Value, Array(Sheets("Справочник").ListObjects("Объекты").Range.Cells(i, 2), Sheets("Справочник").ListObjects("Объекты").Range.Cells(i, 3), Sheets("Справочник").ListObjects("Объекты").Range.Cells(i, 4), Sheets("Справочник").ListObjects("Объекты").Range.Cells(i, 5), Sheets("Справочник").ListObjects("Объекты").Range.Cells(i, 6), coeffSort(i - 1))
        Next i
        'For Each x In objects.Keys
            'Debug.Print x, objects(x)(4)
            'Debug.Print objects.Keys(1)
        'Next x
        'Debug.Print .Item(objects.Keys(1))(0)
        'Debug.Print objects("Волхонка АО " & Chr(34) & "Невский экологический оператор" & Chr(34))(1)
        'Debug.Print objects(objects.Keys(1))(0)
        'Debug.Print objects(Sheets(1).Cells(341, 3).Value)(5)
    End With


    lastRow = Cells(1, 1).CurrentRegion.Cells(Cells(1, 1).CurrentRegion.Cells.Count).Row

    Do While Not Cells(lastRow, 1) = "Итого (масса образования)"
        lastRow = lastRow - 1
    Loop

    Dim zonesCells() As Variant
    Dim zonesAll() As Long
    zonesCells = Range(Cells(4, 1), Cells(lastRow - 1, 1))
    elem = 1
    For e = 1 To UBound(zonesCells) - 1
        If Not zonesCells(e, 1) = zonesCells(e + 1, 1) Then
            ReDim Preserve zonesAll(1 To elem)
            zonesAll(elem) = CLng(zonesCells(e, 1))
            elem = elem + 1
        End If
    Next e

    j = 1
    Dim zones() As Long 'находим количество лотов
    ReDim Preserve zones(1 To 1)
    For Each e In zonesAll
        unique = True
        For elem = LBound(zones) To UBound(zones)
            If zones(elem) = e Then
                unique = False
                Exit For
            End If
        Next elem
        If unique Then
            ReDim Preserve zones(1 To j)
            zones(j) = e
            j = j + 1
        End If
    Next e

    ' For e = 1 To UBound(zones)
        'Debug.Print "zones(e): ", zones(e)
    ' Next e
    'Debug.Print "UBound(zones): ", UBound(zones)

    For Each zone In zones 'удаление листов если уже есть
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = "Лот " & zone Then ws.Delete
        Next ws
    Next zone

    For zone = LBound(zones) To UBound(zones)
        is0 = False
        'ReDim weightsArrObjects(1 To 1), weightsArr0(1 To 1), weightsArr1(1 To 1), kmArr0(1 To 1), kmArr1(1 To 1) As Double
        
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = "Лот " & zone
        Sheets(1).Select
        
        With Sheets("Лот " & zone)
            .Cells(1, 1).Value = Sheets(1).Cells(1, 1).Value
            .Cells(2, 1) = Sheets(1).Cells(2, 1)
            For i = 1 To 11
                .Cells(3, i) = Sheets(1).Cells(3, i)
            Next i

            n = 4
            For i = 4 To lastRow 'заполнение первого блока с МО
                If Sheets(1).Cells(i, 1) = zone Then
                    For j = 1 To 11
                        .Cells(n, j) = Sheets(1).Cells(i, j)
                    Next j
                    n = n + 1
                End If
            Next i

            Dim objectsArr, weightsArr, kmArr, objectsArrFull, weightsArrFull, kmArrFull As Variant

            objectsArr = .Range(.Cells(4, 4), .Cells(n - 1, 4))
            weightsArr = .Range(.Cells(4, 9), .Cells(n - 1, 9))
            kmArr = .Range(.Cells(4, 10), .Cells(n - 1, 10))
            
            objectsArrFull = Sheets(1).Range(Cells(4, 4), Cells(lastRow - 1, 4))
            weightsArrFull = Sheets(1).Range(Cells(4, 9), Cells(lastRow - 1, 9))
            kmArrFull = Sheets(1).Range(Cells(4, 10), Cells(lastRow - 1, 10))

            sumWeightZone = 0
            For e = LBound(weightsArr) To UBound(weightsArr)
                sumWeightZone = sumWeightZone + weightsArr(e, 1) 'масса образования по лоту
            Next e

            .Cells(n, 1) = "Итого"
            .Cells(n, 9) = sumWeightZone
            ReDim Preserve checkingSumWeight(1 To zone)
            checkingSumWeight(zone) = sumWeightZone

            sumMultiplicationResult = 0
            For e = 1 To UBound(weightsArr) 'средневзвешенное расстояние
                multiplicationResult = weightsArr(e, 1) * kmArr(e, 1)
                sumMultiplicationResult = sumMultiplicationResult + multiplicationResult
            Next e
            kmResult = sumMultiplicationResult / sumWeightZone
            .Cells(n, 10) = kmResult 'средневзвешенное расстояние
            .Cells(n, 11) = Application.WorksheetFunction.Sum(.Range(.Cells(4, 11), .Cells(n - 1, 11))) 'сумма РСО
            
            startRow0 = 4 'начальная строка 1 блока без заголовков
            endRow0 = n - 1 'последняяя строка 1 блока без итоговой

            n = n + 1

            ' lastRowNewSh = Cells(Rows.Count, 1).End(xlUp).Row

            Dim sortPlaces() As String
            countObj = 1

            For i = 4 To n 'Наименования объектов прямого вывоза и 1 плеча n = 111
                If Not .Cells(i, 4) = .Cells(i + 1, 4) Then
                    ReDim Preserve sortPlaces(1 To countObj)
                    sortPlaces(countObj) = .Cells(i, 4).Value
                    countObj = countObj + 1
                End If
            Next i

            n = n + 2
            
            startRow1 = n + 2 'начальная строка 2 блока (1 плечо и прямой вывоз итоги) без заголовков
            .Cells(n, 1) = "Первое плечо и прямой вывоз итоги"
            n = n + 1

            Set findCell = Sheets(1).Range(Cells(n, 1), Cells(500, 1)).Find("Первое плечо и прямой вывоз итоги")

            For j = 1 To 13
                .Cells(n, j) = Sheets(1).Cells(findCell.Row + 1, findCell.Column + j - 1)
                .Cells(n, j) = Sheets(1).Cells(findCell.Row + 1, findCell.Column + j - 1)
            Next j

            n = n + 1

            sumWeight0 = 0
            sumWeight1 = 0
            
            startRow = n
            For i = 1 To UBound(sortPlaces)
            
                .Cells(n, 3) = sortPlaces(i)
                
                sumWeightObject = 0
                For row1 = 1 To n 'вес по объекту
                    If .Cells(row1, 4) = .Cells(n, 3) Then
                        sumWeightObject = sumWeightObject + .Cells(row1, 9)
                    End If
                Next row1
                .Cells(n, 5) = sumWeightObject

                .Cells(n, 1) = objects(.Cells(n, 3).Value)(4) 'плечо объекта
                
                result = 0
                For element = 1 To UBound(weightsArr) 'суммпроизв расстояния
                    If objectsArr(element, 1) = .Cells(n, 3).Value Then
                        result = result + (CDbl(weightsArr(element, 1)) * CDbl(kmArr(element, 1)))
                    End If
                Next element
                km = result / sumWeightObject
                .Cells(n, 10) = km
                
                
                If .Cells(n, 1) = "Первое плечо" Then
                
                    .Cells(n, 6) = objects(.Cells(n, 3).Value)(0) 'отбор ВМР
                    '.Cells(n, 7) = objects(.Cells(n, 3).Value)(1) * objects(.Cells(n, 3).Value)(5) 'производственная программа
                    .Cells(n, 7) = .Cells(n, 5) * objects(.Cells(n, 3).Value)(5) 'производственная программа
                    
                    .Cells(n, 8) = CDbl(.Cells(n, 5).Value) - Application.WorksheetFunction.Min(CDbl(.Cells(n, 5).Value), CDbl(.Cells(n, 7).Value)) * CDbl(.Cells(n, 6).Value) 'масса после обраб
                    .Cells(n, 9) = Application.WorksheetFunction.Min(CDbl(.Cells(n, 5).Value), CDbl(.Cells(n, 7).Value)) * CDbl(.Cells(n, 6).Value)
                    
                    

                    If CDbl(.Cells(n, 5)) > CDbl(.Cells(n, 7)) Then
                        .Cells(n, 13) = CDbl(.Cells(n, 5)) - CDbl(.Cells(n, 7)) 'перегруз
                    Else: .Cells(n, 13) = 0
                    End If
                    
                    sumWeight1 = sumWeight1 + CDbl(.Cells(n, 5).Value) 'масса 1 плечо
                    

                ElseIf .Cells(n, 1) = "Прямой вывоз" Then
                    is0 = True

                    .Cells(n, 6) = "—"
                    .Cells(n, 7) = "—"
                    .Cells(n, 8) = "—"
                    .Cells(n, 9) = "—"
                    .Cells(n, 13) = 0 'перегруз
                    
                    sumWeight0 = sumWeight0 + CDbl(.Cells(n, 5).Value) 'масса прямой вывоз

                End If
                
                n = n + 1
            Next i
            endrow = n - 1
            
            Set findCell = .Range(.Cells(startRow, 1), .Cells(endrow, 1)).Find("Первое плечо")
            .Cells(endrow, 12) = sumWeight1 'масса 1 плечо
            If is0 = True Then .Cells(findCell.Row - 1, 12) = sumWeight0 'масса прямого вывоза
            ReDim Preserve checkingObjectsWeight(1 To zone)
            checkingObjectsWeight(zone) = sumWeight1 + sumWeight0 'масса 1 плеча и прямого вывоза
            'Debug.Print checkingObjectsWeight(zone)
                    
            ReDim weightsArr01(1 To endrow - startRow + 1)
            ReDim kmArr01(1 To endrow - startRow + 1)
            ReDim typeArr01(1 To endrow - startRow + 1)
            e = 1
            For i = startRow To endrow
                typeArr01(e) = .Cells(i, 1).Value
                weightsArr01(e) = .Cells(i, 5)
                kmArr01(e) = .Cells(i, 10)
                e = e + 1
            Next i
            
            km1 = 0
            For e = 1 To UBound(typeArr01)
                If typeArr01(e) = "Первое плечо" Then
                    kmWeight = weightsArr01(e) * kmArr01(e)
                    km1 = km1 + kmWeight
                End If
            Next e
            km1 = km1 / sumWeight1
            .Cells(endrow, 11) = km1 'средневзвешенное 1 плечо

            If is0 = True Then
                km0 = 0
                For e = 1 To UBound(typeArr01)
                    If typeArr01(e) = "Прямой вывоз" Then
                        kmWeight = weightsArr01(e) * kmArr01(e)
                        km0 = km0 + kmWeight
                    End If
                Next e
                km0 = km0 / sumWeight0
                .Cells(findCell.Row - 1, 11) = km0 'средневзвешенное прямой вывоз
            End If

            .Cells(n, 1) = "Итого"
            .Cells(n, 5) = sumWeightZone
            .Cells(n, 7) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow, 7), .Cells(endrow, 7)))
            .Cells(n, 8) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow, 8), .Cells(endrow, 8)))
            .Cells(n, 9) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow, 9), .Cells(endrow, 9)))
            .Cells(n, 13) = Application.WorksheetFunction.Sum(.Range(.Cells(startRow, 13), .Cells(endrow, 13)))

            endRow1 = n - 1 'конечная строка 2 блока (1 плечо и прямой вывоз итоги) без итоговой
            
            Dim weightsAfterSort() As Double
            Dim objectsAfterSort() As String
            e = 1
            For i = startRow To endrow
                ReDim Preserve weightsAfterSort(1 To e) 'масса после обработки 1 плечо
                ReDim Preserve objectsAfterSort(1 To e) 'наименования объектов 1 плечо
                If IsNumeric(.Cells(i, 8)) = True Then weightsAfterSort(e) = .Cells(i, 8) Else weightsAfterSort(e) = 0
                objectsAfterSort(e) = .Cells(i, 3)
                e = e + 1
            Next i

            km = 0
            e = 1
            For i = startRow To endrow 'средневзвешенное итог
                kmWeight = weightsArr01(e) * kmArr01(e)
                km = km + kmWeight
                e = e + 1
            Next i
            km = km / sumWeightZone
            .Cells(n, 10) = km

            .Range(.Cells(startRow1, 1), .Cells(endRow1, 13)).Sort Key1:=.Range(.Cells(startRow1, 1), .Cells(endRow1, 1)), Order1:=xlAscending, Header:=xlNo 'сортировка чтобы сначала 1 плечо, потом прямой вывоз

            For i = startRow1 To endRow1 'объединение ячеек расстояний и масс
                If Not .Cells(i, 11) = "" Then
                    .Range(.Cells(startRow1, 11), .Cells(i, 11)).MergeCells = 1
                    .Range(.Cells(startRow1, 12), .Cells(i, 12)).MergeCells = 1
                    mergedRow = i
                    For ii = mergedRow + 1 To endRow1
                        If .Cells(ii, 11) = "" Then
                            .Range(.Cells(ii, 11), .Cells(endRow1, 11)).MergeCells = 1
                            .Range(.Cells(ii, 12), .Cells(endRow1, 12)).MergeCells = 1
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
                For i = startRow To endrow
                    If objects.Keys(e) = .Cells(i, 3) Then sumWeightObjectsZone(e + 1) = .Cells(i, 5)
                Next i

                If sumWeightObjectsZone(e + 1) = Empty Then sumWeightObjectsZone(e + 1) = 0

                For element = LBound(objectsArrFull) To UBound(objectsArrFull)
                    If objectsArrFull(element, 1) = objects.Keys(e) Then sumWeightObjectsFull(e + 1) = sumWeightObjectsFull(e + 1) + weightsArrFull(element, 1)
                Next element
            Next e
        
            '  -----------------------------------2 плечо-----------------------------------------

            startRow2 = endRow1 + 6 'начальная строка 3 блока (2 плечо) без заголовков
            
            Set findCellSheet1 = Sheets(1).Range(Cells(n, 1), Cells(1000, 1)).Find("Второе плечо")
            For j = 1 To 12 'заголовки
                .Cells(startRow2 - 2, j) = Sheets(1).Cells(findCellSheet1.Row, findCellSheet1.Column + j - 1)
                .Cells(startRow2 - 1, j) = Sheets(1).Cells(findCellSheet1.Row + 1, findCellSheet1.Column + j - 1)
            Next j
            startRow = findCellSheet1.Row + 1

            Set findCellSheet1 = Sheets(1).Range(Cells(startRow, 1), Cells(1000, 1)).Find("Итого")
            endrow = findCellSheet1.Row - 1
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
            ReDim Preserve landfills2(1 To (endRow2 - startRow2 + 1))
            ReDim Preserve km2(1 To (endRow2 - startRow2 + 1))
            ReDim Preserve weights2(1 To (endRow2 - startRow2 + 1))
            ReDim Preserve sortedWeights2(1 To (endRow2 - startRow2 + 1))
            ReDim Preserve unsortedWeights2(1 To (endRow2 - startRow2 + 1))

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

            startRow3 = endRow2 + 6 'начальная строка 3 блока (Объекты размещения) без заголовков
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
                For i = startRow1 To endRow1 'цикл по блоку 1 плечо и прямой вывоз итоги
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
            .Range(.Cells(startRow1 - 1, 1), .Cells(endRow1, 2)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow1 - 1, 3), .Cells(endRow1, 4)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow1 - 1, 10), .Cells(startRow1 - 1, 11)).HorizontalAlignment = xlCenterAcrossSelection
            .Range(.Cells(startRow1 - 2, 1), .Cells(startRow1 - 1, 13)).Font.Bold = True
            .Range(.Cells(startRow1 - 2, 1), .Cells(endRow1 + 1, 13)).borders.LineStyle = xlContinuous
            .Range(.Cells(startRow1, 5), .Cells(endRow1 + 1, 5)).NumberFormat = "#,##0.00"
            .Range(.Cells(startRow1, 6), .Cells(endRow1 + 1, 6)).NumberFormat = "0.00%"
            .Range(.Cells(startRow1, 7), .Cells(endRow1 + 1, 13)).NumberFormat = "#,##0.00"
            .Range(.Cells(endRow1 + 1, 1), .Cells(endRow1 + 1, 13)).Font.Bold = True
            .Range(.Cells(endRow1 + 1, 1), .Cells(endRow1 + 1, 4)).HorizontalAlignment = xlCenterAcrossSelection

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
            For i = startRow1 To endRow1 + 1
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
            sort1 = .Range(.Cells(startRow1, 6), .Cells(endRow1, 6))
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
