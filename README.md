# PyVisa
These codes work with "How to program instruments" written by Tao Song in Xi'an Jiaotong-Liverpool University to give a guide on instruments programming based on VISA (Virtual Instrument Software Architecture, it is commonly used in instruments communications).

The terminal numbers of file name indicates the sum of the number of outputs and DMM in your measuring system. VISA- 3 terminals.py is recommended if you need two output channels from power supply to work with DMM.

There are two ways to set up your measurements system. With Keysight (HP, Agilent) instruments, you can check the file "How to program instruments - BenchVue" to get a quick start. With instruments from all brands using Visa, I strongly recommend you to try using a programming way to control your measurements system (see file "How to program instruments - Python"). 

The measurements system contains two parts mainly, one is the controlling of instruments, the second part is set up the probe station. Once you have finished these two parts, it is easy for you to physically connect them together (you can also refer to the file "How to program instruments - BenchVue" for the guide to ser up probe station).
