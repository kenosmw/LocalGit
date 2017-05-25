Attribute VB_Name = "aaClassGlobals"
Option Explicit

'class of products
Public CLS_PRODUCTS As clsProducts

'class of pricing
Public CLS_PRICES As clsPrices

'class of template
Public CLS_TEMPLATES As clsTemplates

'class of userform/jobdetails
Public CLS_USERFORM As clsTemplates

Public Sub testLoadClasses()
    Set CLS_PRODUCTS = New clsProducts
    Set CLS_PRICES = New clsPrices
    
    Call CLS_PRODUCTS.LoadProductDeetsDB
    Call CLS_PRICES.LoadPrices
    
    ProductTesting.Show
    
    Set CLS_PRODUCTS = Nothing
    Set CLS_PRICES = Nothing
End Sub
