Attribute VB_Name = "modSmartBridgeCalculationTypes"
Option Explicit

Public Type BridgeInputs
    SpanM As Double
    BridgeWidthM As Double
    qLive As Double
    qDeck As Double
    deckCost As Double
    dynFactor As Double
    YoungsE As Double
    Fy As Double
    density As Double
    FoS As Double
    deflRatio As Double
    crossSpacingMM As Double
    postSpacingMM As Double
    railHeightMM As Double
    panelLengthMM As Double
    finishMult As Double
    finishR As Long
    finishG As Long
    finishB As Long
End Type

Public Type SectionCalc
    qBeam As Double
    M As Double
    V As Double
    sigma As Double
    tau As Double
    vm As Double
    delta As Double
    f1 As Double
    uB As Double
    uS As Double
    uVM As Double
    uD As Double
    uF As Double
    uMax As Double
End Type

Public Type LayoutCalc
    nCross As Long
    nPosts As Long
    nPanels As Long
    actualCrossSpacing As Double
    totalCost As Double
    totalMass As Double
End Type
