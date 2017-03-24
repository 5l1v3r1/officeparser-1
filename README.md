# officeparser
Python executable to extract VB code from XLS macros

This is a compiled standalone executable of the [officeparser library](https://github.com/unixfreak0037/officeparser/blob/master/officeparser.py) licensed under MIT and copyrighted Copyright (c) 2014 John William Davison.

The executable was built with [PyInstaller](http://www.pyinstaller.org/) on Linux 64bit:
```
$ uname -a
Linux ip-172-30-0-229 4.4.0-66-generic #87-Ubuntu SMP Fri Mar 3 15:29:05 UTC 2017 x86_64 x86_64 x86_64 GNU/Linux
$ python --version
Python 2.7.12
$ pyinstaller -F --ascii --strip officeparser.py
...
```
