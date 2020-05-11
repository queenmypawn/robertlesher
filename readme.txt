The `path` variable defines where the mystencils folder containing all of
the imported stencils will be created.
This is found in line 27 and defines the usage of lines 27-36.

The `columns` variable defines the standardized columns of
the stanardized .csv file that is created.
This is found in line 37 and defines the usage of lines 37-50.

The `newStencil` variable on line 65 defines the imported stencil.
Wherever your stencil is, you should change this variable
to the path of the stencil, located in the mystencils folder.

When you want to enter the device information onto the .csv,
the command `input(...)` on line 129 requests that you press the `Enter`
or `Return` button on your keyboard when you have saved this information.

Lines 125 and 126 define the start setting of the initial shape.
It is currently set to drop the collection into the middle
of the board. To change the initial shape's location, change `x` and `y`.

Lines 130 and 134 output the .csv files onto the working directory.

Line 145 determines that the devices will be printed in columns of 4
to keep the Visio drawing readable. If you would like to print the objects
in a different setting, the `devicesInCol` variable should be changed.

Line 146 determines the spacing between columns in page units.

Line 151 determines the text used to describe the dropped shape at
the fourth parameter of dropShape.

Line 164 determines the text to append to the connection between shapes.