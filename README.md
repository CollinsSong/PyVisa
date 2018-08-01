# PyVisa
These codes will work together with "How to program instruments" written by Tao Song in Xi'an Jiaotong-Liverpool University to give a guide on instruments programming based on VISA (Virtual Instrument Software Architecture, it is commonly used in instruments communications).

Files: ID-VG precise, ID-VD precise and Gate leakage precise are for MOSFET testing. You can run it directly.

The file VISA Communications Testing is used to check if you have a nice connection between your PC and instruments. 

There are two ways to set up your measurements system. With Keysight (HP, Agilent) instruments, you can check the file "How to program instruments - BenchVue" to get a quick start of graphical programming. With instruments from all brands using Visa, I strongly recommend you to try to use programming way to control your instruments (see file "How to program instruments - Python"). 

The measurements system contains two parts mainly, one is the controlling of instruments, the second part is the setting up of probe station. Once you have finished these two parts, it is easy for you to connect them together physically (you can also refer to the file "How to program instruments - BenchVue" for the guide to set up probe station).
