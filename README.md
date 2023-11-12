# Trinomial_tree
Implementation of a trinomial tree pricing model in VBA and Python

This code, coded in classes both in Python and VBA replicates the trinomial pricing model for equity options that distributes dividend

The trinomial pricing model is as follows : 

$$
dS_t = S_t r dt + S_t \sigma dW_t - D_t
$$

With Dt a discrete dividend and the ex-div date is a parameter of the pricing model;

We use the following assumptions:
- Time steps are all equal: Δt
- The middle node is equal to the forward price
- Nodes values are geometric series: α = Si,j+1/Si,j
- The next middle node is the closer to the forward price
- α is defined by a multiple of the standard deviation over one time step ≈ St σ sqr(Δt):  Si,j+1 - Si,j ≈ sqr(3) StdDev
- Divide by Si,j: α ≈ 1+ sqr(3) StdDev / Si,j
- The actual formula is: 

$$ 
𝛼 = 𝑒xp(𝑟Δ𝑡+𝜎3Δ𝑡) 
$$

For each node we do the following : 
- The probabilities are specified so as to ensure that the price of the underlying evolves as a martingale, while the moments – considering node spacing and probabilities – are matched to those of the log-normal distribution

- The first moment : 

$$ 
E_{i+1,j} = p_{up} S_{i+1,j'+1} + p_{mid} S_{i+1,j'} + p_{down} S_{i+1,j'-1} = E[S_{t,i+1} | S_{t,i}]

$$ 