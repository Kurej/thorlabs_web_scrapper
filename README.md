# Thorlabs web scrapper



## Description

This repository intends to gather weights from Thorlabs references or from a Thorlabs cart.





## Usage

To gather weights from  Thorlabs shopping cart exported as a file named `shoppingCart.xls` placed in this directory, just run `thorlabs_scrapper.py`.

To gather the weight of a specific component, for instance the reference [MBT616D/M](https://www.thorlabs.com/thorproduct.cfm?partnumber=MBT616D/M) you can run the following code snippet:

```python
import thorlabs_scrapper as ts
ts.get_thorlabs_product_weight('MBT616D/M')
```

It returns the following result:

`MBT616D/M weighs 1.29 kg`





## Installation

Navigate to the code directory and create a virtual environment named `.venv` located in `<path>/thorlabs_scrapper/.venv/` using:

`> python -m venv .venv`

Activate the virtual environment:

`> & .venv/Scripts/Activate.ps1`

Install necessary packages:

`(.venv)> pip install -r requirements.txt`