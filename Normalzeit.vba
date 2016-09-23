Function Normalzeit(dblIZeit As Single) As String
REM Normalzeit.vba - Diese Funktion berechnet die Normalzeit im Format TT:SS:MM, ausgehend
REM von der Industriezeit AZES
REM
REM Copyright (c) 2016 Pasqualino Imbemba p.imbemba@gmail.com
REM
REM This program is free software: you can redistribute it and/or modify
REM it under the terms of the GNU General Public License as published by
REM the Free Software Foundation, either version 3 of the License, or
REM (at your option) any later version.
REM 
REM This program is distributed in the hope that it will be useful,
REM but WITHOUT ANY WARRANTY; without even the implied warranty of
REM MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
REM GNU General Public License for more details.
REM
REM You should have received a copy of the GNU General Public License
REM along with this program.  If not, see <http://www.gnu.org/licenses/>.

    Dim intTage As Integer
    Dim intIMin As Integer 'Industrieminuten
    Dim intNMin As Integer 'Normale Minuten
    Dim strTage As String
    Dim strStunden As String
    Dim strMinuten As String
    Dim arrHun() As String
    Dim singArbeitstag As Single
    Dim strTmpMin As String
    
    singArbeitstag = 7.6
    
    
    If dblIZeit > 0 Then
        intTage = 0
        While dblIZeit >= 7.6
            dblIZeit = Round(dblIZeit, 2)
            intTage = intTage + 1
            dblIZeit = dblIZeit - singArbeitstag
        Wend
        
        'Tage
        strTage = CStr(intTage)
        strTage = AddZeroFirst(strTage)
        Normalzeit = strTage & ":"
        
        If dblIZeit > 0 Then
        
            'Stunden
            arrHun = Split(CStr(dblIZeit), ",")
            strStunden = CStr(arrHun(0))
            strStunden = AddZeroFirst(strStunden)
            Normalzeit = Normalzeit & strStunden & ":"
        
            Rem Minuten. Umrechnung siehe https://www.fin.be.ch/fin/de/index/personal/anstellungsbedingungen/arbeitszeit/sollarbeitszeit.assetref/dam/documents/FIN/PA/de/Minuten-Umrechnungstabelle%20d.pdf
            strMinuten = "00"
            
            If UBound(arrHun()) > 0 Then
                strTmpMin = Left(arrHun(1), 2)
                If Len(strTmpMin) = 1 Then
                    strTmpMin = strTmpMin & "0"
                End If
                
                intIMin = CInt(strTmpMin)
                intNMin = (intIMin / 5) * 3
                strMinuten = CStr(intNMin)
                strMinuten = AddZeroFirst(strMinuten)
            End If
    
            Normalzeit = Normalzeit & strMinuten
        Else
            Normalzeit = Normalzeit & "00:00"
        End If
    Else
        Normalzeit = "00:00:00"

    End If

End Function


Private Function AddZeroFirst(strString As String) As String
    If Len(strString) = 1 Then
        AddZeroFirst = "0" & strString
    Else
        AddZeroFirst = strString
    End If
End Function
