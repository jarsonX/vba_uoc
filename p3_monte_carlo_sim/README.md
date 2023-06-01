## Problem statement

A great tool for investigating uncertainty in a complex process is the Monte Carlo simulation.  If you'd like, there is plenty of information online that describes what this simulation does and what it is useful for.  But in brief, a Monte Carlo simulation aims to simulate the possible pathway of a future endeavor, experiment, or process, given multiple inputs that each have uncertainty.  Monte Carlo simulations are oftentimes used by financial planners to try to predict what might happen in the future.  Inputs to financial planning processes include probabilities of how the stock market might do, projections for various costs and amount of sales, and many other financial variables.  In the sciences, there may be several or many inputs into a future process, and the goal is to understand the likelihood of a scenario going one way or another.

In this project, you will first learn about a Monte Carlo simulation and how to implement it in a VBA user form using an example investigating a cookie recipe, where each of the ingredients has an uncertainty associated with it.  You will explore five different typical distributions used in Monte Carlo simulations and implement these into the VBA user form.  Finally, you will adapt the cookie example into the main project deliverable, which is a user form that allows the user to simulate a profitability analysis based on net present value (NPV) of a proposed capital project.

## Requirements

1. User is able to input bounds for the various distributions as indicated in the problem statement. When the GO button is clicked, the simulation works (does not crash) for at least 1,000 simulations.

2. The percentage of simulations that resulted in positive net present value (NPV) is output in a message box.

3. Accuracy. When inputting a set of values for various distributions, the simulation should output a result that is close to the solution (+/- 2%).

4. The results of the simulation are output successfully in a histogram (based on the code provided by the starter files).

