<div align="center">

## Frequency Analysis


</div>

### Description

Learn how to break cyphertext with Frequency Analysis.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2005-01-02 14:29:44
**By**             |[Daniel M](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-m.md)
**Level**          |Intermediate
**User Rating**    |3.5 (21 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Frequency\_183583122005\.zip](https://github.com/Planet-Source-Code/daniel-m-frequency-analysis__1-58066/archive/master.zip)





### Source Code

After reading "The Code Book", by Simon Singh I have been inspired to start cryptanalysis and have been attempting to write strong encryption algorithms and break them. This article focuses on how to use Frequency Analysis - <b>the method of determining substituted characters by analyzing the frequency, or repetition/iterations of characters and comparing them to standard English.</b><P>
First, I will give a brief overview of the steps taken to use Frequency Analysis.<p>
1. The first step of Frequency Analysis is to count up the frequencies of each character in the ciphertext. I have included a .zip which automatically does this for you. There should be about five letters in which have a frequency less than 1% and they are most likely the letters j, k, q, x, and z. One of the letters should have a frequency greater than 10%, which probably represents the letter "e". That is, this generalization occurs only if the language it is written in follows the frequency chart for it's specific language, in this case, is English.<p>
2. If the frequency chart follows the english frequency chart but decipherment is still not possible, the next step is to focus on pairs of repeated letters. For instance, in English, the most commonly repeated letters are as follows: ss, ee, tt, ff, ll, mm, and oo. If the ciphertext has any repeated characters, you can assume that they are one of those.<p>
3. If the ciphertext has spaces between words, then try to decipher words that contain a length of less than four letters. Here is a list of one to three letter words that are most common and can be tried when deciphering:<br>
1 Letter: A, I<br>
2 Letters: of, to, in, it, is, be, as, at, so, we, he, by, or, on, do, if, me, my, up, an, go, no, us, am<br>
3 Letters: the, and<p>
4. If it is possible, find english texts that are similar to the ciphertext and use those for your frequency chart to get a most accurate chart. For instance, excerpt taken from "The Code Book"<p>
<i>"military messages tend to omit pronouns and articles, and the loss of words such as <b>I</b>, <b>he</b>, <b>a</b> and <b>the</b> will reduce the frequency of some of the commonest letters. If you know you are tackling a military message, you should use a frequency table generated from other military messages."</i><p>
5. A skill commonly used in frequency analysis is the ability to indentify words or whole phrases based on experience or guesses. For example, if the military sends an encrypted weather report at 6:00 PM everyday, you can possibly assume the first word of the ciphertext may be the word "Weather", in which you could use to help break the rest of the ciphertext. These are known as <b>cribs</b>.<p>
6. Last, but not least, if two frequency charts seem to match, but the ciphertext is not readable, this draws the conclusion that the text is indeed not a substitution cipher, but a transposition cipher.<P>
<b>7.</b> Further methods of frequency analysis become more complicating, but can further help a person break a cipher. These include gathering statistics on the relationships between letters -how often a letter is seen neighboring another letter, or how often a letter begins a new word or ends a new word. Frequency analysis is a powerful tool for deciphering text if you follow the correct steps.
<P>
The .zip I have included offers a frequency chart generator in which you can export the information to a .txt file. I plan to further this project into a fully-functional Frequency Analysis Decrypter utility which will take the user step-by-step to decrypt the text. Thank you for reading.

