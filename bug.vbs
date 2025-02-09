Late Binding and Type Mismatches: VBScript uses late binding by default, meaning variable types aren't checked until runtime. This can lead to unexpected errors if a variable is used in a way inconsistent with its actual type.  For example, trying to perform arithmetic on a string that can't be converted to a number will result in an error.

Example:
```vbscript
dimension myVar
myVar = "10a"
result = myVar + 5
'Error: Type mismatch
```