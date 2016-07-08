#pptconv.net
Windows C# .NET solution for converting PowerPoint documents to PDF using LibreOffice.

##Requirements
Before you can use `pptconv.exe`, you need to have [LibreOffice](https://www.libreoffice.org/) installed. If you install it in your `Program Files`, no further configuration is required.

If you install it somewhere other than `Program Files`, you can define the `PPTCONV_LIBREOFFICE_PATH` environment variable to point to where you installed LibreOffice.

##Usage
To convert slideshows `slide1.pptx`, `slide2.pptx`, `slide3.pptx`, and so on:

```BASH
pptconv.exe slide1.pptx slide2.pptx slide3.pptx ...
```

Absolute and relative paths are acceptable. Inputs are converted one at a time and output will be placed in your current working directory.

##API Usage
Most classes are self-documented. The main Converter class is named `LibreOfficeConverter`. Two main APIs are `Queue` and `Flush`.

`Queue` Adds files to be converted in the background converter thread, and `Flush` waits (blocks the calling thread) for the pending files to be converted.

`LibreOfficeConverter` propagates two events: `ConversionSucceed` and `ConversionFailed` with a path to the input file.

Additionally you can use `LibreOfficeProcess` class to use the core converter in your own custom queuing system.

##Building From Source
Open up `pptconv/pptconv.sln` and do a `Clean/Rebuild`. Additionally you may use [Test Explorer](https://msdn.microsoft.com/en-us/library/hh270865.aspx) to run all unit tests. Resulting assembly would be `libpptconv.dll`.
