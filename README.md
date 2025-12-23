# neXt Voice Engine (VB.NET)

A concatenative speech engine based on phoneme stitching with crossfading.

## Features

- Phoneme Parser: Recognizes complex sounds like "sh," "ei," and "ch."
- Diphone Support: Prioritizes transitions (e.g., "ba," "scha") if they exist as files.
- Audio Stitching: Joins WAV files in memory.
- Crossfading: Prevents clicks and pops with smooth transitions.

## Installation

1. Include the `SiriVoiceEngine.dll` as a reference in your project.
2. Create a folder for your `.wav` files.
3. The files must be named exactly like the sound (e.g., `sh.wav`, `a.wav`, `ba.wav`).


## Example (VB.NET)
```vbnet
Dim engine = New neXtVoiceEngine("C:\Path\To\Sounds\", rules)
engine.Speak("Hello World")
