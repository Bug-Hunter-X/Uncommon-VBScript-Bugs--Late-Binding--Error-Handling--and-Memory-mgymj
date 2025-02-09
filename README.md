# Uncommon VBScript Bugs and Solutions

This repository explores some less frequently encountered errors in VBScript programming, focusing on late binding issues, improper error handling with `On Error Resume Next`, and potential memory management problems in larger scripts.

The `bug.vbs` file demonstrates these issues, while `bugSolution.vbs` provides corrected versions with explanations.

## Bugs Addressed:

* **Late Binding and Type Mismatches:** VBScript's dynamic typing can lead to runtime errors if types are not carefully managed.
* **`On Error Resume Next` Misuse:**  Over-reliance on this statement can mask serious errors and make debugging harder.
* **Memory Management in Larger Scripts:**  VBScript's garbage collection might not always be efficient, particularly with large datasets or complex object usage.