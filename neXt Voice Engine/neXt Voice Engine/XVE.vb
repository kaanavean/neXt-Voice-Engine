Imports System.IO
Imports System.Linq
Imports System.Collections.Generic
Imports System.Media

''' <summary>
''' A class for concatenating audio data to achieve speech output.
''' </summary>
Public Class XVE
    Private _vocalPath As String
    Private _rules As New List(Of String)

    ''' <summary>
    ''' Creates a new instance of the Synthesizer module.
    ''' </summary>
    ''' <param name="wavFolder">The folder containing the required .wav phonemes.</param>
    Public Sub New(ByVal wavFolder As String)
        ' Check and ensure that the specified path ends with a backslash.
        _vocalPath = If(wavFolder.EndsWith("\"), wavFolder, wavFolder & "\")

        ' Automatically creates a list of possible phonemic rules
        Dim allFiles = Directory.GetFiles(_vocalPath, "*.wav")
        For Each file In allFiles
            Dim name As String = Path.GetFileNameWithoutExtension(file)
            _rules.Add(name)
        Next
        ' Sort the rules by the number of letters, from largest to smallest
        _rules = _rules.OrderByDescending(Function(s) s.Length).ToList()
    End Sub

    ''' <summary>
    ''' Processing the text and speaking the combined audio data.
    ''' </summary>
    Public Sub Speak(ByVal text As String)
        If String.IsNullOrWhiteSpace(text) Then Return

        ' Break down text into phonemes
        Dim phonemes = ParseToPhonemes(text)

        ' Merge audio data with the file of the crossfade module in one memory location
        Dim combinedWave = CombineWithCrossfade(phonemes, 441) ' 441 samples, the equivalent of 10 ms

        ' Play the combined audio data
        If combinedWave IsNot Nothing AndAlso combinedWave.Length > 44 Then
            Try
                Using ms As New MemoryStream(combinedWave)
                    Using player As New SoundPlayer(ms)
                        player.PlaySync()
                    End Using
                End Using
            Catch ex As Exception
                Debug.WriteLine("Fehler bei der Wiedergabe: " & ex.Message)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Breaks down text into phoneme abbreviations based on the established rules.
    ''' </summary>
    Private Function ParseToPhonemes(ByVal input As String) As List(Of String)
        Dim result As New List(Of String)()
        Dim text As String = input.ToLower().Replace(" ", "_") ' An underscore corresponds to a space, for the pause WAVs
        Dim i As Integer = 0

        While i < text.Length
            Dim found As Boolean = False

            ' Checks if Diphones are present. Five characters are requested in advance; more are possible.
            For length As Integer = 5 To 2 Step -1
                If i + length <= text.Length Then
                    Dim candidate As String = text.Substring(i, length)
                    If File.Exists(Path.Combine(_vocalPath, candidate & ".wav")) Then
                        result.Add(candidate)
                        i += length
                        found = True
                        Exit For
                    End If
                End If
            Next

            ' Automatic error correction, repeating step 1, can be removed in case of high latency
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

            ' Simple letter
            If Not found Then
                Dim c As String = text(i).ToString()
                If File.Exists(Path.Combine(_vocalPath, c & ".wav")) Then
                    result.Add(c)
                End If
                i += 1
            End If
        End While
        Return result
    End Function
    ''' <summary>
    ''' Reads the WAV files and combines them into a valid WAV stream.
    ''' </summary>
    Private Function CombineWithCrossfade(ByVal phonemes As List(Of String), ByVal fadeSamples As Integer) As Byte()
        Dim audioParts As New List(Of Short())
        Dim header As Byte() = Nothing

        ' The path to the start tone. Connection to the dictionary is being developed
        Dim startSoundPath As String = "C:\XDev\neXt Voice Synthesizer\end_recognition.wav"

        ' Load the start sound as the first element
        Dim allFiles As New List(Of String)()

        ' The existence of the start sound is checked and included at the beginning of the stream
        If File.Exists(startSoundPath) Then
            allFiles.Add(startSoundPath)
        End If

        ' Adding the phonemes
        For Each p In phonemes
            allFiles.Add(Path.Combine(_vocalPath, p & ".wav"))
        Next

        ' Convert each file into short arrays
        For Each filePath In allFiles
            If File.Exists(filePath) Then
                Try
                    Dim fileBytes = File.ReadAllBytes(filePath)
                    If fileBytes.Length > 44 Then
                        ' Extracting the header from the start sound file
                        If header Is Nothing Then header = fileBytes.Take(44).ToArray()

                        ' Extract audio-data
                        Dim rawData = fileBytes.Skip(44).ToArray()
                        Dim shorts(rawData.Length / 2 - 1) As Short
                        Buffer.BlockCopy(rawData, 0, shorts, 0, rawData.Length)
                        audioParts.Add(shorts)
                    End If
                Catch ex As Exception
                    Debug.WriteLine("Fehler beim Laden von: " & filePath)
                End Try
            End If
        Next

        If audioParts.Count = 0 Then Return Nothing

        ' Calculate crossfade
        Dim resultList As New List(Of Short)()

        For i As Integer = 0 To audioParts.Count - 1
            Dim currentPart = audioParts(i)

            If i = 0 Then
                ' Add the first sample (the start sound) completely
                resultList.AddRange(currentPart)
            Else
                ' Crossfade between the end of the previous stream and the new part
                Dim overlapStart = resultList.Count - fadeSamples
                If overlapStart < 0 Then overlapStart = 0

                ' Crossfade math
                For s As Integer = 0 To fadeSamples - 1
                    If s < currentPart.Length And (overlapStart + s) < resultList.Count Then
                        Dim fadeOutFactor As Double = 1.0 - (s / fadeSamples)
                        Dim fadeInFactor As Double = s / fadeSamples

                        Dim mixedSample = (resultList(overlapStart + s) * fadeOutFactor) + (currentPart(s) * fadeInFactor)
                        resultList(overlapStart + s) = CShort(mixedSample)
                    End If
                Next

                ' Attach the rest
                If currentPart.Length > fadeSamples Then
                    resultList.AddRange(currentPart.Skip(fadeSamples))
                End If
            End If
        Next

        ' Final Byte Array Creation
        Dim finalBytes(resultList.Count * 2 + 43) As Byte
        Buffer.BlockCopy(header, 0, finalBytes, 0, 44)
        Buffer.BlockCopy(resultList.ToArray(), 0, finalBytes, 44, resultList.Count * 2)

        ' WAV Header Correction
        Using ms As New MemoryStream(finalBytes)
            Using writer As New BinaryWriter(ms)
                ms.Position = 4
                writer.Write(CInt(ms.Length - 8))
                ms.Position = 40
                writer.Write(CInt(resultList.Count * 2))
            End Using
            Return ms.ToArray()
        End Using
    End Function
End Class