# BrewCalibCollect
Automating the collection of calibration files in MaximDL. 

Maxim provides the facility to set calibration files, so images are immediately calibrated (Bias, Dark, Flat) when they are taken.

The process involves:
- Collecting the raw frames, for example 
    - 25 flats using Red filter and 1x1 bin -10 degC, 
    - 40 Bias frames at 1x1 bin and -10 degC, 
    - 25 Dark frames at 1x1 bin and -10 DegC.
- Delete the existing Master files matching these frames
- In Maxim, go into the Set Calibration dialog
- Autogenerate. This brings in the new files as individual files.
- Replace with Masters. This combines the frames (i.e., 25 red flats) into a single Master frame
- Exit Maxim
- Delete the individual raw frames

This is kind of a lot of steps, with some possibility of error. It would be nice to automate this process, perhaps with an ACP script or similar.

Unfortunately Maxim does not provide scripting access to all of the functions. I think there is an autogenerate type function, but no Replace with Masters. This program is intended to do these tasks.
