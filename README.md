# Active Works To Google Docs Spreadsheet
A simple Python script that reads the registration report CSV file from Active
Works and parses the information into a Google Docs Spreadsheet as used by
http://www.abdcycling.com/ Indoor TT Series for registration and results.

## Dependencies
- Python >= 2.7
- [Google Data Python Client Library](http://code.google.com/p/gdata-python-client)
- [Python Progressbar](http://pypi.python.org/pypi/progressbar)

## Use
    $ ./aw2gd.py --user [google username] --pw [google password] --input [CSV file]

- username: your gmail username
- password: your gmail password (if you have special characters, place password in "quotes")
- input: location of the local CSV file from Active Works

## Authors
- Richard A. Johnson

## License
(The GNU General Public License)

Copyright &copy; 2011 Richard A. Johnson &lt;nixternal@gmail.com&gt;

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
