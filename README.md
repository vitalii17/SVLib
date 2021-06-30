# SVLib
Altium Library created and maintained by Vitalii Shunkov

##Project Structure
###Template project structure:
Project name
|-- Board
|   |-- Project Name + Version
|       |-- Source
|       |-- Documentation
|           |-- Scheematic
|           |-- Assembly
|           |-- BOM
|       |-- Gerbers and Drill
|           |-- Single
|           |-- Panel
|       |-- Pick and Place
|       |-- 3D Model
|
|
|-- Case
|   |-- Drawings
|   |-- STL
|   |-- 3D Design
|
|
|-- Calculations
|
|
|-- .git


##Symbol library
Power dissipation for SMD chip resistors:
5W     - V   - 2512(6332), ...
2W     - ||  - 2512(6332)
1W     - |   - 2010(5025), 2512(6332)
0.5W   - __  - 2010(5025)
0.25W  - \   - 1206(3216)
0.125W - \\  - 0603(1608), 0805(2012), 1206(3216)
0.05W  - \\\ - 0201(0603), 0402(1005)


##Footprint library
Courtyard indent:
0.125 for component size less or equal 0402
0.185 for component size equal         0603 (easy to use in 50mil grid for 1.27mm pitch ICs)
0.200 for componnents equal            0805
0.250 for componnents more or equal    1206

How to write footprint name:
QFN: QFN<Pitch>P<Body Size X>X<Body Size Y>X<Body Height>-<Leads Number>TP (Example: QFN40P600X600X90-48TP)
QFP: QFP<Pitch>P<Body Size X>X<Body Size Y>X<Body Height>-<Leads Number>   (Example: QFP50P1000X1000X100-64N)
Fuse Resettable SMD: FUSERC<Body Length + Body Width>(<Inch Body Length + Inch Body Width>)X<Body Height>
PLS: HDRV<Pin Number>P<Pin Pitch>_<Pin Per Row Number>X<Row Number>_W<Lead Width>_<Body Length>X<Body Thickness>X<Component Height>

How to write description:
<Component Type>, <Leads Number>, <Body Size>
Example: Chip Capacitor, Body 0.6x0.3mm

Switches types:
SPST - Single Pole Single Throw Switch
SPDT - Single Pole Double Throw Switch
DPST - Double Pole Single Throw Switch
DPDT - Double Pole Double Throw Switch
PBS  - Push Button Switch
TS   - Toggle Switch
LS   - Limit Switch
FS   - Float Switches
FLS  - Flow Switches
PS   - Pressure Switches
TS   - Temperature Switches
JS   - Joystick Switch
RS   - Rotary Switches

Naming suffixes:
F  - Flat Leads  (Example: DIOM120X80X63_F)
TP - Thermal Pad (Example: QFN40P600X600X90-48TP)
R  - Right Angle

##3D Components library
3D body colors       - (R,   G,   B):
Metal Leads          - (210, 209, 199)
Plastic case (black) - (37, 36, 36)
First  pin marking   - (176, 169, 152)

##Datasheet library
###File naming:
<Manufacturer> - <Part Number>
<Manufacturer> - <Part Number xxxx>, where xxxx - Part Number pattern
<Manufacturer> - <Part Number 1>_<Part Number 2>
<Manufacturer> - <Part Serie>
<Manufacturer> - <Part Serie 1>_<Part Serie 2>

Example: "NXP - MMA8452Q.pdf"
         "Murata - BLM18HxxxxSN1x.pdf"
         "Coils Electronic - CCSP_CCFH_HCFT.pdf"

##Layers naming convention
Mechanical 1 : Assembly Top
Mechanical 2 : Assembly Bottom
Mechanical 3 : 3D Body Top
Mechanical 4 : 3D Body Bottom
Mechanical 5 : Coutyard Top
Mechanical 6 : Coutyard Bottom
Mechanical 7 : Component Center Top
Mechanical 8 : Component Center Bottom
Mechanical 9 : Designator Top
Mechanical 10: Designator Bottom
Mechanical 11: Board Outline
Mechanical 12: Mill-Panel
Mechanical 13: V-Cut
Mechanical 14: Mill-Board
Mechanical 16: Dimensions-Board
Mechanical 17: Dimensions-Panel
Mechanical 18: Geometry Restrictions Top
Mechanical 19: Geometry Restrictions Bottom
Mechanical 20: Consruction Geometry
Mechanical 21: Stencil Fiducials Top
Mechanical 22: Stencil Fiducials Bottom

##Installation
Install DBLib:
1. Install "./Install/AccessDatabaseEngine_X64.exe".
2. Go to "Components" -> "Hamburger button" -> "File-based Libraries Preferences...". Go to "Installed" tab. Press "Install".
3. Install Microsoft Access (optional, install if you want to modify the library).

Install Templates:
1. Go to "Settings/Data Management/Templates". In "Template Location" select path: "./Template".
2. Go to "Settings/Draftsman/Templates". In "Templates Location" select path: "./Template".
3. Go to "Settings/System/New Document Defaults". Select defaults for PCB, Schematic, Drftsman and OutputJob files.

Load settings:
1. Go to "Preferencess". Go to "Load..." -> "Load from file...". Select "./Preferences/DXP_Preferences.DXPPrf". 
On "Options Page" deselect:
    - "Data Management/File-based Libraries". Press "OK".
    - "System/Installation"
    - "Data Management/Templates"
    - "Draftsman/Templates"

##Misc
EIA  - Metric:
0201 - 0603
0402 - 1005                 
0603 - 1608                  
0805 - 2012                    
1206 - 3216                  
1210 - 3225             
2010 - 5025                  
2512 - 6332  

DO-214 packages:
DO-214AA - SMB, middle size
DO-214AB - SMC, largest size
DO-214AC - SMA, smallest size
DO-214BA - GF1

SOD packages:
SOD-523 - 120X80mm
SOD-323 - 170X125mm
SOD-123 - 275X160mm

SOT package - JEDEC Name:
SOT-23-3    - TO-236
SOT-143     - TO-253
SOT-23-5    - MO-178 (AA)
SOT-23-6    - MO-178 (AB)
SOT-23-8    - MO-178 (BA)

SOIC Name - Standard Name 
SOxx-150  - JEDEC MS-012
SOxx-300  - JEDEC MS-013
SOxx-208  - EIAJ EDR-732

Density Level:
L - Least, or minimum copper
M - Most, or maximum copper
N - Nominal, or median copper   

Operating temperature:
Commercial:   0 °C to +70 °C
Industrial: -40 °C to +85 °C
Military  : -55 °C to +125 °C

##ToDo
- Convert dblib to svndblib.
- Move each footprint and symbol in separate files for better VCS support.
        

