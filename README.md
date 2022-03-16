# Discrete Probability Detector in VB6

Discrete Probability Detector (DPD) is an algorithm that transforms any sequence of symbols into a transition matrix. The algorithm may receive special characters from the entire ASCII range. These characters can be letters, numbers or special characters (ie. <kbd>`q#7Eu9f$*"</kbd>). The number of symbol/character types that make up a string, represent the number of states in a Markov chain. Thus, DPD is able to detect the number of states from the sequence and calculate the transition probabilities between these states. The final result of the algorithm is represented by a transition matrix (square matrix) which contains the transition probabilities between these symbol types (or states). The transition matrix can be further used for different prediction methods, such as Markov chains or Hidden Markov Models. This version of DPD is made in Visual Basic 6.0.

# Screenshot

![screenshot](https://github.com/Gagniuc/Discrete-Probability-Detector-in-VB6/blob/main/screenshot/DPD%20(1).PNG)
![screenshot](https://github.com/Gagniuc/Discrete-Probability-Detector-in-VB6/blob/main/screenshot/DPD%20(2).PNG)

# References

- <i>Paul A. Gagniuc. Markov chains: from theory to implementation and experimentation. Hoboken, NJ,  John Wiley & Sons, USA, 2017, ISBN: 978-1-119-38755-8.</i>
