Hedging Vanilla Options — VBA Implementation
=============================================

Excel/VBA project implementing dynamic delta hedging and static replication
of a Call Up-and-Out (CUO) barrier option.

The workbook simulates hedging strategies for a CUO and benchmarks their P&L
distributions across multiple spot/maturity scenarios.

MODULES
-------

BaseFunctions.bas
  Core Black-Scholes library:
  - myN(x)            Standard normal CDF
  - mySwitch()        Maps "Call"/"Put" to +1/-1
  - myBS()            Black-Scholes price (call or put, with dividends)
  - myDelta()         BS Delta
  - myGamma()         BS Gamma
  - myTheta()         BS Theta
  - myVega()          BS Vega
  - myRho()           BS Rho
  - myCRR()           Binomial tree pricer (Cox-Ross-Rubinstein)

  CUOHedge.bas
  Analytical pricing and static replication of a Call Up-and-Out:
  - cuo_fran()        Closed-form CUO price (reflection principle)
  - numdelta_cuo()    Numerical delta via finite differences
  - Stat_hedge_cuo()  Static replication using a strip of vanillas at the barrier
Scenarios.bas

  Monte Carlo simulation of the dynamic delta hedge:
  - myHedge_CUO()     Simulates GBM paths, applies delta hedging at each step,
                      records hedge error (knocked-out vs. survived)
  - Scenarios_UOC()   Loops over a 5x5 grid of spot levels (85-115) and
                      maturities (0.1-1Y), outputs results and summary statistics
    ScenarioChart.bas
  Visualization:
  - CreateChartSheet()  Creates the UOC_Chart sheet with selectable S and T inputs
  - RefreshChart()      Plots histogram of normalised hedge costs for the selected
                        scenario; highlights static hedge cost in red

ThisWorkbook.cls
  Sets Excel calculation to manual on open (required for Monte Carlo performance)

Chart1.cls
  Reminds the user to press F9 to trigger a new simulation trajectory

KEY CONCEPTS
------------

  CUO closed-form      Reflection principle (Rubinstein & Reiner)
  Delta hedge          Numerical delta recomputed at each time step
  Static replication   Strip of vanillas at barrier (Carr & Chou)
  Hedge cost metric    NormCost = HedgeError / UOC_Price
  Knock-out detection  Path-wise barrier check at each GBM step

USAGE
-----

1. Open the .xlsm in Excel with macros enabled
2. Set parameters on the dynDeltaHedge sheet (K, r, mu, sigma, q, H, nbTraj, nbSteps)
3. Run Scenarios_UOC to launch the full 25-scenario simulation
4. Navigate to UOC_Chart, select a scenario (S, T) to visualise the hedge cost distribution
