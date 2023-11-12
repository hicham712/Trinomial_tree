# Trinomial_tree
Implementation of a trinomial tree pricing model in VBA and Python

This code, coded in classes both in Python and VBA replicates the trinomial pricing model for equity options that distributes dividend

The trinomial pricing model is as follows : 

$$
dS_t = S_t r dt + S_t \sigma dW_t - D_t
$$

We use the following assumptions:
- Time steps are all equal: Δt
- The middle node is equal to the forward price
- Nodes values are geometric series: α = Si,j+1/Si,j
- The next middle node is the closer to the forward price