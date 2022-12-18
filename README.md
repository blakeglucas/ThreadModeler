# ThreadModeler

CLI Version of the ThreadModeler Software offered by coolOrange on the [Cloud](https://threadmodeler.com/) and as an [Autodesk Extension](https://apps.autodesk.com/INVNTOR/en/Detail/Index?id=2540506896683021779). Original source code can be found [here](https://github.com/coolOrangeLabs/inventor-thread-modeler).

## Building/Running

This project relies on .NET CLI version 6.0.200. Running the command

```
dotnet restore
```

will install the required dependencies.

This application also relies on Autodesk Inventor. Ensure you have a valid, licensed version of Autodesk Inventor running before starting this application, as it will not start Inventor automatically. 

This application has only been tested with Autodesk Inventor 2023. Other versions may/may not work out of the box, but the logic is sound.

## Copyright Notices

As reproduced in code, `Toolkit.cs` and `Worker.cs` provide the following notices
```
Copyright (c) Autodesk, Inc. All rights reserved
Written by Philippe Leefsma 2011 - ADN/Developer Technical Services

This software is provided as is, without any warranty that it will work. You choose to use this tool at your own risk.

Neither Autodesk nor the author Philippe Leefsma can be taken as responsible for any damage this tool can cause to your data. Please always make a back up of your data prior to use this tool, as it will modify the documents involved in the feature transformation.
```

`Program.cs` includes code taken from [`Controller.cs`](https://github.com/coolOrangeLabs/inventor-thread-modeler/blob/master/ThreadModeler/Commands/Controller.cs) in [this repo](https://github.com/coolOrangeLabs/inventor-thread-modeler).