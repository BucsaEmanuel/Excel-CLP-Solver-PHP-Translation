# Excel-CLP-Solver-PHP-Translation
Attempt at translating the Excel CLP Solver (original author Güneş Erdoğan)

Original file can be found at: https://people.bath.ac.uk/ge277/index.php/clp-spreadsheet-solver/

I extracted the two VB scripts and started the tedious process of translation to PHP.

Since I don't know any visual basic, I expect this to take a long time.

You might notice that most lines carry comments above. I use phpStorm and add docblocks along with the original VB code.
e.g.
<pre>
/**
 * Dim max_z As Double
 * @var float
 */
$max_z = 0.0;
</pre>

The original code that's being translated in this example is <pre>Dim max_z As Double</pre>