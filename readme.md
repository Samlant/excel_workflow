 <!-- title: IntakeTool -->

![Logo](https://i.postimg.cc/GmPY2yGJ/E6-B76-AE1-3407-4-BB7-8-D7-F-159613-FC9-C7-D.png)

This tool expedites and simplifies intaking new or renewal clients by automatically organizing their file, adding them to our tracker, and---optionally---allows you to allocate markets (and submit to them) in one fell swoop.

# Table of Contents <!-- omit in toc -->

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
  - [Start the Program:](#start-the-program)
  - [Using the Program](#using-the-program)
- [Screenshots](#screenshots)
- [FAQ](#faq)
  - [How can I tell it's working?](#how-can-i-tell-its-working)
  - [Can I change which folder I can add new qouteforms to?](#can-i-change-which-folder-i-can-add-new-qouteforms-to)
- [Support](#support)
- [License](#license)
- [Authors](#authors)
- [Related](#related)

# Features

| ▪️  | Feature List                                                                                        |
| --- | --------------------------------------------------------------------------------------------------- |
| 🥂  | Fully cross-platform.                                                                               |
| 🥂  | automatically detects if the quoteform is for a new client or a renewal.                            |
| 🥂  | Automatically creates an entry in our excel tracking report.                                        |
| 🥂  | Allows you to allocate markets right off the bat---if so, it includes those markets in the tracker. |
| 🥂  | Also allows you to immeidately submit the client into markets, if desired.                          |
| 🥂  | Displays an icon within the System Tray in the bottom-right corner                                  |

# Installation

Run the setup file.

- This will create the program in:

```
C:/Users/{username}/AppData/Local/work-tools
```

# Usage

## Start the Program:

---

First things first, the program starts upon computer boot-up, so no manual start-up is required.

## Using the Program

---

In order to trigger the program, simply save a new quoteform into the FL Office shared drive (don't save into any specific folder, just the base of the drive).

The program will react automatically; once the program recognizes a new quoteform (can be either a Word doc or PDF quoteform), the program will show a popup for you to select which action you would like to choose for this specific client at this time. The options include:

- Create a folder for the client and add to the tracker, nothing further;

- Create a folder and allocate markets but don't submit to the markets at this time (a new dialog will appear showing checkboxes next to each market; select each option that you would like to submit this client to);

- Create a folder, allocate markets and submit to them as well, all in one fell swoop (in this case, a new program will launch, but will carry over all information necessary from the created quoteform such as client name & vessel details. You will simply need to select which markets that you would like to submit to, and, optionally, attach any other files or insert any additional text into the emails to underwriters).

- After clicking your desired option, the program will execute without further notice, so you are free to continue with your work as normal.

# Screenshots

![App Screenshot](https://via.placeholder.com/468x300?text=App+Screenshot+Here)

# FAQ

## How can I tell it's working?

You'll see an icon in the system tray icon in the lower-right corner. Also, when you add a file to the watched folder, a dialogue window will open.

## Can I change which folder I can add new qouteforms to?

No, not currently. We are planning to be able to add this eature by right-clicking on the system tray icon in the lower-right corner. For now, you may request it from us and we'll adjust it accordingly for you.

# Support

For support or any questions, email sam@novamar.net

# License

[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](https://choosealicense.com/licenses/mit/)

# Authors

- [@Samlant](https://github.com/Samlant)

# Related

Here are some related projects to this one:

[QuickDraw: A tool to quickly submit a client into many markets using just their quoteform](https://github.com/Samlant/IntakeTool)
