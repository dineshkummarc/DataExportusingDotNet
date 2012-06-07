'@ ---------------------------------------------------------
'@ Prashant Nayak's Export Buttons common mudule
'@ Copyright (C) 2004 Prashant Nayak. All rights reserved.
'@ Code change history:
'@ ~~~~~~~~~~~~~~~~~~~~
'@ 1. Bug fix - 12/07/2004 
'@    Use Delimiters in HEADERs.
'@ ---------------------------------------------------------
Option Strict On
Option Explicit On 

Imports System.Text
Namespace PNayak.Web.UI.WebControls
    Module Common
        '@<summary>
        '@ Convert dataview to TAB, Comma or any character separated string.
        '@</summary>
        Public Function ConvertDataViewToString(ByVal srcDataView As DataView, Optional ByVal Delimiter As String = Nothing, Optional ByVal Separator As String = ",") As String

            Dim ResultBuilder As StringBuilder
            ResultBuilder = New StringBuilder()
            ResultBuilder.Length = 0

            Dim aCol As DataColumn
            For Each aCol In srcDataView.Table.Columns
                If Not Delimiter Is Nothing AndAlso (Delimiter.Trim.Length > 0) Then
                    ResultBuilder.Append(Delimiter)
                End If
                ResultBuilder.Append(aCol.ColumnName)
                If Not Delimiter Is Nothing AndAlso (Delimiter.Trim.Length > 0) Then
                    ResultBuilder.Append(Delimiter)
                End If
                ResultBuilder.Append(Separator)
            Next
            If ResultBuilder.Length > Separator.Trim.Length Then
                ResultBuilder.Length = ResultBuilder.Length - Separator.Trim.Length
            End If
            ResultBuilder.Append(Environment.NewLine)

            Dim aRow As DataRowView
            For Each aRow In srcDataView
                For Each aCol In srcDataView.Table.Columns
                    If Not Delimiter Is Nothing AndAlso (Delimiter.Trim.Length > 0) Then
                        ResultBuilder.Append(Delimiter)
                    End If
                    ResultBuilder.Append(aRow(aCol.ColumnName))
                    If Not Delimiter Is Nothing AndAlso (Delimiter.Trim.Length > 0) Then
                        ResultBuilder.Append(Delimiter)
                    End If
                    ResultBuilder.Append(Separator)
                Next aCol
                ResultBuilder.Length = ResultBuilder.Length - 1
                ResultBuilder.Append(vbNewLine)
            Next aRow
            If Not ResultBuilder Is Nothing Then
                Return ResultBuilder.ToString()
            Else
                Return String.Empty
            End If
        End Function
    End Module
End Namespace
