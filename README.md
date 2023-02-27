# STREQ
A GUI for calculating structural equivalencies of matrices (Euclidean Distance, Jaccard Matching, Simple Matching).

![alt text](https://github.com/mbiggiero/STREQ/blob/main/screenshot.png?raw=true)

Input: 

Excel matrices. Values on diagonal are ignored.
Excel file with a single sheet automatically calculates intra-matrix distances/matchings, 2 sheets for inter-matrices distance/matching.
Errors thrown when the matrix isn't square or when matrices in the 2 sheets have different labels.


Output:

ED/JM/SM Inter: single .txt file with Absolute and Normalized results (additional Column/Row results for ED only);

ED/JM/SM Intra: single .xls file with 6 sheets (Total/Row/Column x Normalized/Absolute)

