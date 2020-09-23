<div align="center">

## Code Optimization Tips


</div>

### Description

This article contains some optimization tips for both standard coding and game creation.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jay \(God\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jay-god.md)
**Level**          |Intermediate
**User Rating**    |3.8 (46 globes from 12 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jay-god-code-optimization-tips__1-30210/archive/master.zip)





### Source Code

<FONT NAME="Arial" SIZE="2">
   Some responses to this article have brought things I may have missed to my attention, so #8, #10, and all numbers beyond 16 are the tips that have been added or modified. Thank you to everyone who added their constructive comments!<P>
   As the computing industry rapidly increases speed and ability of hardware (especially processors), programmers have mainly found a leisure in writing quality code. Visual Basic is the noteable example.<P>
     There are several ways to absolutely optimize your Visual Basic source, and I will demonstrate some actual processor speed differences between different commands (they are really only calculable from very old systems).<P>
     All of my experience comes from taking calculations from a Pentium I 75 MHz machine with 40 MB RAM. Whether you are writing a game or creating an intensive looping algorithm, these tips will definitely help save memory and speed up your application.<P>
<B>#1. ELSEIF VS. SELECT CASE</B><BR>
     While the select case has provided easier readability, especially as switch in C/C++, it really does nothing more than put a series of ELSEIF's into a cleaner format. The difference is very very neglible in most cases, SELECT CASE being only about 3% slower, but in an algorithm this could make a major difference. If you need some additional speed, replace your SELECT CASE with IF/ELSEIF/ELSE ("Default:" being the equivalent of ELSE).<P>
<B>#2. Data Types</B><BR>
     The Data Types used in Visual Basic have an adverse effect on both speed AND memory space. One Data Type I will never suggest using in Visual Basic is the Boolean type, which takes up two bytes of memory for a True or False operand. If you are using solely true and false and do not need numerals over 255 to be defined as true, try using a Byte type, which allows numbers from 0 to 255. Simply specify to yourself that 0 is false and 1 is true. Byte is both faster and smaller than Boolean and will definitely make a difference in game programming.<P>
<B>#3. Data Types II</B><BR>
     The Variant Data Type is a very poor type to use on a consistent basis. If you are able to define a Data Type, do so. The difference will usually be around 20 bytes of memory freed up. *NEVER* use an array of Variants unless it is completely unavoidable.<BR>
     IMPORTANT NOTE: For game programmers, the difference of selecting the proper Data Type is very important. Please take into consideration that the integer Data Type is two bytes and the Byte type is one. However, the integer Data Type is actually accessed faster than byte by a fraction of a fraction of a nanosecond.<P>
<B>#4. Data Types III</B><BR>
     As a final note on Data Types, you can define the length of your strings in order to conserve memory. This is done with an asterisk as follows during definition:<P>
     Dim strData as String * 15<P>
     This will define the size of the string as fifteen characters. Strings use up the most memory space of any data type, requiring memory slots for data information as well as the actual array of bytes that make up the string. By defining your string as 15 characters, it will only take up 15 memory slots. However, if you do not define it as fixed length, it will take up 25 memory slots (STRING SIZE + 10).<BR>
     Finally, data types that use decimals will always take longer to perform the mathematical equations to calculate the decimal, so avoid them unless they are needed for your application.<P>
<B>#5. Arrays</B><BR>
     Arrays can be very useful, but sometimes are larger than their counterparts would be. An array takes up 20 bytes of memory plus 4 bytes per dimension. Take this into account before making a three dimensional array that will hold only three pieces of data. Is the overall sizes of the added data types going to justify using the array? Usually the array will justify itself with ease of use in an algorithm.<BR>
     If you finish with a public array and no longer need it, destroy it to conserve memory.<P>
<B>#6. Write Nice Code</B><BR>
     Self explanatory. Use algorithms, not millions of if statements, etc.<P>
<B>#7. Functions</B><BR>
     Functions are indeed the most important aspect of programming. However, calling a function does slow down the processor by a few billionths of a second. If you have a function that draws a line with the LineTo API, for example, don't call the function but call the API directly from your loop. You'll definitely notice a difference.<P>
<B>#8. Using the Windows Drawing Functions</B><BR>
     Instead of using the pset, line, circle, point and arc functions, replace them with the Windows API equivalents (GetPoint, SetPixel, LineTo, etc). SetPixel is 75% faster than PSet, and LineTo is 50% faster than Line (Yes, I have calculated to exact billionths of a measurement). If you are using an algorithm (movement for example) then definitely use the APIs.<BR>
   Instead of using SetPixel, use SetPixelV. SetPixel returns a value to VB containing the Long value of the color used, which is superfluous. SetPixelV is not only faster, but it does not require 4 bytes of memory to be stored in a return value each time.<BR>
   The only problem with this immense speed gaining method is that VB6.0 contains a bug. If you modify your DC (IE: PictureBox) with the WinAPI drawing functions, you usually have to do a Picture.Refresh to show the changes. I suppose no compiler is perfect.<P>
<B>#9. LOOPING</B><BR>
     The fastest loop is the FOR loop, and it should be used whenever possible. The slowest loop is the WHILE WEND loop, and DO comes in a close second behind the FOR loop.<BR>
     Also, when you are using a FOR loop, do not specify the variable in the NEXT statement. A good practice is to comment out the variable, to avoid confusion. For example:<P>
     For i = 1 to 32656<BR>
          'INTENSE CODE<BR>
     Next 'i<P>
     This will speed up your loops, and is especially suggested when nesting is in practice.<P>
<B>#10. GOTO/GO SUB</B><BR>
     The worst programming method is definitely the GOTO command. All functions in a program's code should lead back to an original location, so to speak. The GOTO actually breaks this by making the next line of execution a random label. In the breakdown of the code, this means slower execution and possibly a larger executable file. Avoid GOTO at all costs (it is always possible).<BR>
   In response to using ON ERROR GOTO LABELNAME, this is a very shoddy method of error handling that should be used sparingly under necessary circumstances, or during design time testing. To avoid having to use this, consider using Windows API to do error-prone functions (errors are usually set by the functions, and you can see if an error occurred by checking the return value or by calling GetLastError). For those of you who don't know, there is an API in shell32.dll called SHFileExist that tests if a file exists or not; using VB error checking to check for this is an example of code methods NOT to use. I'll also point out that your Error is automatically a variant data type, so by declaring ON ERROR you have added over 20 bytes of memory as soon as the sub routine is called. This will usually be more than all of your declared data type sizes combined!<P>
<B>#11. Knowing when to use EXIT</B><BR>
     Exit Sub, Exit Function, Exit For, and Exit Select are all useful for ending a function or routine prematurely. If you have a large series of if statements, and only one can be true, then Exit Sub after it has completed its portion of code. This will make the compiler skip over the rest of the code, which would be skipped anyway. The result is slightly faster execution (unless the chosen statement is the last!)<P>
<B>#12. Constants</B><BR>
     Avoid using constants whenever possibly if you need speed. Instead of specifying the constant, which needs to be referenced from memory, simply give the compiler the number or string directly from the design code. This probably goes for VB constants as well (assuming VB does not include them if they are not used).<P>
<B>#13. Private VS. Public</B><BR>
     If you have a function or Data Type that will only be used by one form or module, declare it as private. This will release the memory after it has completed, while a public and static variable will remain in memory until your program completes. Private is much faster and should definitely be used when necessary (including Declarations of APIs).<P>
<B>#14. Load</B><BR>
     A neglible difference can be seen by effective use of Sub Main() instead of beginning execution with a form. This also gives you more control over the application's starting process (and whether you need to dynamically create windows or controls on load).<P>
<B>#15. Unload</B><BR>
     Make sure that you delete your brushes and DCs and clear your images and modifications with Control.Property = Nothing<BR>
     Obviously, don't forget to release all of your forms, processes, and threads before ending execution.<P>
<B>#16. Picture Use [Skinning]</B><BR>
     If you must skin your software, use an internal resource and BitBlt the pieces. Skinning will definitely give slower computers a problem with speed and unless necessary, keep it to a minimum.<P><BR>
<B>#17. Object Properties and Variables</B><BR>
   Never store your variables in objects. Do not make your textbox say "true" or "false" depending on your code, store this in a variable.<BR>
   It takes VB from 5 to 15 times the length to get the data from an object property than from a variable. With this in mind, if you intend on looping or using a property repeatedly, then store the property in a variable and call the variable instead of the property. For example, you would change:<P>
   for i = 1 to 1000<BR>
     X = Picture1.Width + i<BR>
   next 'i<BR>
to:<P>
   J = Picture1.Width<BR>
   for i = 1 to 1000<BR>
     X = J + i<BR>
   next 'i<P>
<B>#18. Using Drawing Functions</B><BR>
   When you are using the windows API drawing functions, it is best to know which to use under what circumstances. For example, if LineTo has a speed of .58 and SetPixel has a speed of .22, then you would use SetPixel to draw lines of two pixels and LineTo to draw lines of more pixels. In the two pixel example, SetPixel would take .44 seconds while LineTo would take .58. However, if the line was 30 pixels, then LineTo would take around .58 and SetPixel would take 6.6 seconds (A vastly higher time frame!) In most situations where you need a line of a single color drawn, use LineTo over SetPixel.<P>
     That's all for now. <B>PLEASE READ THE DATA TYPE SUMMARY IN YOUR HELP FILE TO SEE THE SIZES AND RANGES OF ALL DATA TYPES</B>, and if I have forgotten something please notify me [jay@necrocosm.com]. Remember: just because you don't notice any speed problems, you may have users on 486 systems who will be very unhappy.<BR></FONT>

