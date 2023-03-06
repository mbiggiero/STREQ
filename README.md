# STREQ - STRuctural EQuivalence 
A program for calculating the following forms of structural equivalence of matrice: Euclidean Distance, Jaccard Matching, Simple Matching.  
These forms of structural equivalence can be calculated for binary and weighted networks, and within a single matrix or between two matrices.   
A deepening on these methods and references to network analysis literature, and a discussion on its applications to inter-firm networks can be found in the Methodological Appendix of the book "Inter-firm Networks", written by Lucio Biggiero & Robert Magnuszewski for the Springer Series on Relational Economics and Organizational Governance. It can be downloaded for free (by clicking on Back Matter) from https://link.springer.com/book/10.1007/978-3-031-17389-9

![screenshot](https://github.com/mbiggiero/STREQ/blob/main/screenshot.png?raw=true)

Input:  
Excel file with a single sheet automatically calculates intra-matrix distances/matchings, 2 sheets for inter-matrices distance/matching.  
Errors thrown when the matrix isn't square or when matrices in the 2 sheets have different labels.  
Note: values on diagonal are ignored.  


Output:  
ED/JM/SM Inter: single .txt file with Absolute and Normalized results (additional Column/Row results for ED only)  
ED/JM/SM Intra: single .xls file with 6 sheets (Total/Row/Column x Normalized/Absolute)

