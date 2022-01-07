# Rocket BlueZone API UDF

## Table of Contents
+ [About](#about)
+ [Prerequisites](#prerequisites)
+ [Installation](#installation)
+ [Documentation](#documentation)

## About <a name = "about"></a>
This UDF is an AutoIt wrapper for 'Rocket Bluezone' 3270 terminal emulator.   
It uses exposed COM objects to communicate to emulator on application level.  
```Note, that not all functionality might be mapped or updated in current published version.```

## Prerequisites <a name = "prerequisites"></a>
User must have 'Rocket Bluezone' application installed and COM objects exposed.

## Installation <a name = "installation"></a>
Simply copy ```RocketBZ_*.au3``` files to your development directory and use ```#include``` to point to these files in the source code.  

## Documentation <a name = "documentation"></a>
* [RocketBZ COM object](https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_acon_bz-host-automation-object.htm)
* [RocketBZ API methods](https://www3.rocketsoftware.com/bluezone/help/v71/en/bzsh/bzaa/source/bzaa_aref_bz-object-methods.htm)