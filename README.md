# Weighted-Kappa-Excel-Function
Add-In for Excel which allows calculating a Weighted Kappa between two raters

The following helped me to understand how the weighting works, and test my function:
https://en.wikipedia.org/wiki/Cohen's_kappa
http://vassarstats.net/kappa.html

My wkap() function allows the passing of two parameters

  =wkap(param1,param2)

param1 should be the Observed Scores Matrix
param2 should be the Weight Matrix (optional param)
  if param2 is not passed, a generic Weight Matrix will be generated internally
    0 for the main diagonal, and incrementing 1 for each cell off of the main diagonal
          
The function generates an Expected Scores Matrix
and calculates the Weighted Kappa by comparing each Score Matrix to the Weighted Matrix
