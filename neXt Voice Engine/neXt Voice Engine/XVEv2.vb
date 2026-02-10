Imports System.IO
Imports System.Linq
Imports System.Collections.Generic
Imports System.Media

Public Class XVEv2
    Private _basePath As String
    Private _monotonePath As String
    Private _stressPath As String
    Private _isAutomatic As Boolean = False
    Private _rules As New List(Of String)

    Public Sub New(ByVal wavFolder As String)
        _basePath = If(wavFolder.EndsWith("\"), wavFolder, wavFolder & "\")

        Dim cleanPath As String = _basePath.TrimEnd(Path.DirectorySeparatorChar)
        Dim folderName As String = Path.GetFileName(cleanPath).ToLower()

        If folderName.EndsWith(".axsm") Then
            _isAutomatic = True
            _monotonePath = Path.Combine(_basePath, "Monotone.arc\")
            _stressPath = Path.Combine(_basePath, "Stress.arc\")
        ElseIf folderName.EndsWith(".sxsm") Then
            _monotonePath = _basePath
            _stressPath = _basePath
        Else
            _monotonePath = _basePath
            _stressPath = _basePath
        End If

        LoadRules(_monotonePath)
    End Sub

    Private Sub LoadRules(path As String)
        If Directory.Exists(path) Then
            Dim allFiles = Directory.GetFiles(path, "*.wav")
            For Each file In allFiles
                Dim name As String = System.IO.Path.GetFileNameWithoutExtension(file)
                If Not _rules.Contains(name) Then _rules.Add(name)
            Next
            _rules = _rules.OrderByDescending(Function(s) s.Length).ToList()
        End If
    End Sub

    Private Structure PhonemeRequest
        Public Token As String
        Public UseStress As Boolean
    End Structure

    Public Sub Speak(ByVal text As String)
        If String.IsNullOrWhiteSpace(text) Then Return
        Dim sequence = BuildPhonemeSequence(text)
        Dim combinedWave = CombineWithCrossfade(sequence, 441)

        If combinedWave IsNot Nothing Then
            Try
                Using ms As New MemoryStream(combinedWave)
                    Using player As New SoundPlayer(ms)
                        player.PlaySync()
                    End Using
                End Using
            Catch ex As Exception
                Debug.WriteLine("Fehler: " & ex.Message)
            End Try
        End If
    End Sub

    Private Function BuildPhonemeSequence(ByVal input As String) As List(Of PhonemeRequest)
        Dim sequence As New List(Of PhonemeRequest)()
        Dim words As String() = input.Split(" "c)
        Dim nextWordStressed As Boolean = False

        For Each word In words
            Dim currentWord As String = word
            Dim stressThisWord As Boolean = False

            If currentWord.StartsWith("!") AndAlso currentWord.EndsWith("!") Then
                stressThisWord = True
                currentWord = currentWord.Trim("!"c)
            End If

            If _isAutomatic AndAlso nextWordStressed Then
                stressThisWord = True
                nextWordStressed = False
            End If

            If currentWord.EndsWith(",") Then nextWordStressed = True

            Dim phonemes = ParseToPhonemes(currentWord)
            For Each p In phonemes
                sequence.Add(New PhonemeRequest With {.Token = p, .UseStress = stressThisWord})
            Next
            sequence.Add(New PhonemeRequest With {.Token = "_", .UseStress = False})
        Next
        Return sequence
    End Function

    Private Function ParseToPhonemes(ByVal input As String) As List(Of String)
        Dim result As New List(Of String)()
        Dim text As String = input.ToLower()
        Dim i As Integer = 0
        While i < text.Length
            Dim found As Boolean = False
            For length As Integer = 5 To 2 Step -1
                If i + length <= text.Length Then
                    Dim candidate As String = text.Substring(i, length)
                    If _rules.Contains(candidate) Then
                        result.Add(candidate)
                        i += length
                        found = True
                        Exit For
                    End If
                End If
            Next
            If Not found Then
                For Each rule In _rules
                    If i + rule.Length <= text.Length Then
                        If text.Substring(i, rule.Length) = rule Then
                            result.Add(rule)
                            i += rule.Length
                            found = True
                            Exit For
                        End If
                    End If
                Next
            End If
            If Not found Then
                result.Add(text(i).ToString())
                i += 1
            End If
        End While
        Return result
    End Function

    Private Function CombineWithCrossfade(ByVal sequence As List(Of PhonemeRequest), ByVal fadeSamples As Integer) As Byte()
        Dim audioParts As New List(Of Short())
        Dim header As Byte() = Nothing
        Dim startSoundPath As String = "C:\XDev\neXt Voice Synthesizer\end_recognition.wav"

        If File.Exists(startSoundPath) Then
            Dim startData = LoadWavToShorts(startSoundPath, header)
            If startData IsNot Nothing Then audioParts.Add(startData)
        End If

        For Each req In sequence
            Dim targetPath As String = If(req.UseStress, _stressPath, _monotonePath)
            Dim filePath As String = Path.Combine(targetPath, req.Token & ".wav")

            If Not File.Exists(filePath) Then filePath = Path.Combine(_monotonePath, req.Token & ".wav")

            If File.Exists(filePath) Then
                Dim data = LoadWavToShorts(filePath, header)
                If data IsNot Nothing Then audioParts.Add(data)
            End If
        Next

        If audioParts.Count = 0 OrElse header Is Nothing Then Return Nothing

        Dim resultList As New List(Of Short)()
        For i As Integer = 0 To audioParts.Count - 1
            Dim currentPart = audioParts(i)
            If i = 0 Then
                resultList.AddRange(currentPart)
            Else
                Dim overlapStart = Math.Max(0, resultList.Count - fadeSamples)
                Dim actualFade = Math.Min(fadeSamples, currentPart.Length)

                For s As Integer = 0 To actualFade - 1
                    Dim fadeOutFactor As Double = 1.0 - (s / actualFade)
                    Dim fadeInFactor As Double = s / actualFade
                    Dim mixed = (resultList(overlapStart + s) * fadeOutFactor) + (currentPart(s) * fadeInFactor)
                    resultList(overlapStart + s) = CShort(mixed)
                Next
                If currentPart.Length > actualFade Then
                    resultList.AddRange(currentPart.Skip(actualFade))
                End If
            End If
        Next

        Dim finalBytes(resultList.Count * 2 + 43) As Byte
        Buffer.BlockCopy(header, 0, finalBytes, 0, 44)
        Buffer.BlockCopy(resultList.ToArray(), 0, finalBytes, 44, resultList.Count * 2)

        Using ms As New MemoryStream(finalBytes)
            Using writer As New BinaryWriter(ms)
                ms.Position = 4 : writer.Write(CInt(ms.Length - 8))
                ms.Position = 40 : writer.Write(CInt(resultList.Count * 2))
            End Using
            Return ms.ToArray()
        End Using
    End Function

    Private Function LoadWavToShorts(path As String, ByRef header As Byte()) As Short()
        Try
            Dim fileBytes = File.ReadAllBytes(path)
            If fileBytes.Length > 44 Then
                If header Is Nothing Then header = fileBytes.Take(44).ToArray()
                Dim rawData = fileBytes.Skip(44).ToArray()
                Dim shorts(rawData.Length / 2 - 1) As Short
                Buffer.BlockCopy(rawData, 0, shorts, 0, rawData.Length)
                Return shorts
            End If
        Catch : End Try
        Return Nothing
    End Function
End Class