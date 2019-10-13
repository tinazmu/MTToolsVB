'On Nuget Console type Install-Package ExcelDna.AddIn
Imports ExcelDna.Integration

Public Module MTTools
    Public Structure ratesStruct
        Public EffectiveDate As Date
        Public ExpiryDate As Date
        Public Rate As Decimal

        Public Sub New(f As Date, t As Date, r As Decimal)
            EffectiveDate = f
            ExpiryDate = t
            Rate = r
        End Sub
    End Structure

    ReadOnly Rates() As ratesStruct = New ratesStruct() {
            New ratesStruct(#2000-01-01#, #2000-06-30#, 11.0),
            New ratesStruct(#2000-07-01#, #2000-12-31#, 12.0),
            New ratesStruct(#2001-01-01#, #2001-06-30#, 12.25),
            New ratesStruct(#2001-07-01#, #2001-12-31#, 11.0),
            New ratesStruct(#2002-01-01#, #2002-06-30#, 10.25),
            New ratesStruct(#2002-07-01#, #2002-12-31#, 10.75),
            New ratesStruct(#2003-01-01#, #2003-06-30#, 10.75),
            New ratesStruct(#2003-07-01#, #2003-12-31#, 10.75),
            New ratesStruct(#2004-01-01#, #2004-06-30#, 11.25),
            New ratesStruct(#2004-07-01#, #2004-12-31#, 11.25),
            New ratesStruct(#2005-01-01#, #2005-06-30#, 11.25),
            New ratesStruct(#2005-07-01#, #2005-12-31#, 11.5),
            New ratesStruct(#2006-01-01#, #2006-06-30#, 11.5),
            New ratesStruct(#2006-07-01#, #2006-12-31#, 11.75),
            New ratesStruct(#2007-01-01#, #2007-06-30#, 12.25),
            New ratesStruct(#2007-07-01#, #2007-12-31#, 12.25),
            New ratesStruct(#2008-01-01#, #2008-06-30#, 12.75),
            New ratesStruct(#2008-07-01#, #2008-12-31#, 13.25),
            New ratesStruct(#2009-01-01#, #2009-06-30#, 10.25),
            New ratesStruct(#2009-07-01#, #2009-12-31#, 9.0),
            New ratesStruct(#2010-01-01#, #2010-06-30#, 9.75),
            New ratesStruct(#2010-07-01#, #2010-12-31#, 10.5),
            New ratesStruct(#2011-01-01#, #2011-06-30#, 10.75),
            New ratesStruct(#2011-07-01#, #2011-12-31#, 10.75),
            New ratesStruct(#2012-01-01#, #2012-06-30#, 10.25),
            New ratesStruct(#2012-07-01#, #2012-12-31#, 9.5),
            New ratesStruct(#2013-01-01#, #2013-06-30#, 9.0),
            New ratesStruct(#2013-07-01#, #2013-12-31#, 8.75),
            New ratesStruct(#2014-01-01#, #2014-06-30#, 8.5),
            New ratesStruct(#2014-07-01#, #2014-12-31#, 8.5),
            New ratesStruct(#2015-01-01#, #2015-06-30#, 8.5),
            New ratesStruct(#2015-07-01#, #2015-12-31#, 8.0),
            New ratesStruct(#2016-01-01#, #2016-06-30#, 8.0),
            New ratesStruct(#2016-07-01#, #2016-12-31#, 7.75),
            New ratesStruct(#2017-01-01#, #2017-06-30#, 7.5),
            New ratesStruct(#2017-07-01#, #2017-12-31#, 7.5),
            New ratesStruct(#2018-01-01#, #2018-06-30#, 7.5),
            New ratesStruct(#2018-07-01#, #2018-12-31#, 7.5),
            New ratesStruct(#2019-01-01#, #2019-06-30#, 7.5),
            New ratesStruct(#2019-07-01#, #2019-12-31#, 7.25)
    }

    <ExcelFunction(Name:="CALCINTFCOA", Description:="Calculates Simple Interest Using FCoA Post Judgment Rates (until 31-12-2019)")>
    Public Function CalcIntMyWay(
                            <ExcelArgument(Name:="Base Value", Description:="is the principal on which the interest will be calculated")>
                            BaseValue As Decimal,
                            <ExcelArgument(Name:="From Date", Description:="is the date the interest calculation starts (inclusive)")>
                            FromDate As Date,
                            <ExcelArgument(Name:="To Date", Description:="is the date the interest calculation ends (inclusive)")>
                            ToDate As Date) As Object
        If FromDate > ToDate Then
            Return ExcelError.ExcelErrorValue
        ElseIf ToDate > Rates(UBound(Rates, 1)).ExpiryDate Or FromDate < Rates(LBound(Rates, 1)).EffectiveDate Then
            Return ExcelError.ExcelErrorNA
        Else
            Dim thisRate As ratesStruct, retVal As Decimal = 0, Days As Integer
            For Each thisRate In Rates
                If thisRate.EffectiveDate > ToDate Or thisRate.ExpiryDate < FromDate Then 'do nothing; there is no overlap

                Else
                    Days = DateDiff(DateInterval.Day, If(thisRate.EffectiveDate < FromDate, FromDate, thisRate.EffectiveDate),
                                    If(thisRate.ExpiryDate < ToDate, thisRate.ExpiryDate, ToDate)) + 1
                    retVal = retVal + (BaseValue * Days * thisRate.Rate / 100.0 / (365.0 - 28.0 + System.DateTime.DaysInMonth(Year(thisRate.EffectiveDate), 2)))
                End If
                'Console.WriteLine(thisRate.EffectiveDate &"."& Days & "." & )
            Next thisRate
            Return retVal
        End If
    End Function

End Module
