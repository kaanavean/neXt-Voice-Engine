# neXt Voice Engine (VB.NET)

A concatenative speech engine based on phoneme stitching with crossfading (or simply put, a synthesizer).

## Features

- Phoneme Parser: Recognizes complex sounds like "sh," "ei," and "ch."
- Diphone Support: Prioritizes transitions (e.g., "ba," "scha") if they exist as files.
- Audio Stitching: Joins WAV files in memory.
- Crossfading: Prevents clicks and pops with smooth transitions.

## Installation

1. Include the `neXt Voice Engine.dll` as a reference in your project.
2. Create a folder for your `.wav` files.
3. The files must be named exactly like the sound (e.g., `sh.wav`, `a.wav`, `ba.wav`).

## Current Status

Currently, all possible phoneme combinations for the English language are being collected in order to create a sound bank.

A stress system is currently being designed that will automatically stress certain words, such as those following a comma. A marking system is also planned that will then assign these words a synthetic stress (in other words, a random stress pattern to make words stand out, etc.).

## Example (VB.NET)
```vbnet
Dim engine = New XVE("C:\Path\To\Sounds\")
engine.Speak("Hello World")
