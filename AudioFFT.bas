Attribute VB_Name = "AudioFFT"
Option Explicit

'These don't change in this program, so I made them constants so they're
'as fast as can be.
Public Const AngleNumerator = 6.283185   ' 2 * Pi = 2 * 3.14159265358979
Public Const NumSamples = 1024
Public Const NumBits = 10

'Used to store pre-calculated values
Private ReversedBits(0 To NumSamples - 1) As Long

Sub DoReverse()
    'I pre-calculate all these values.  It's a lot faster to just read them from an
    'array than it is to calculate 1024 of them every time FFTAudio() gets called.
    Dim I As Long
    For I = LBound(ReversedBits) To UBound(ReversedBits)
        ReversedBits(I) = ReverseBits(I, NumBits)
    Next
End Sub

Function ReverseBits(ByVal Index As Long, NumBits As Byte) As Long
    Dim I As Byte, Rev As Long
    
    For I = 0 To NumBits - 1
        Rev = (Rev * 2) Or (Index And 1)
        Index = Index \ 2
    Next
    
    ReverseBits = Rev
End Function

Sub FFTAudio(RealIn() As Integer, RealOut() As Single)
    'In this case, NumSamples isn't included (since it's always the same),
    'and the imaginary components are left out since they have no meaning here.
    
    'I've used Singles instead of Doubles pretty much everywhere.  I think this
    'makes it faster, but due to type conversion, it actually might not.  I should
    'check, but I haven't.
    
    'The imaginary components have no meaning in this application.  I just left out
    'the parts of the calculation that need the imaginary input values (which is a
    'big speed improvement right there), but we still need the output array because
    'it's used in the calculation.  It's static so that it doesn't get reallocated.
    Static ImagOut(0 To NumSamples - 1) As Single
    
    'In fact... I declare everything as static!  They all get initialized elsewhere,
    'and Staticing them saves from wasting time reallocating and takes pressure off
    'the heap.
    Static I As Long, j As Long, k As Long, n As Long, BlockSize As Long, BlockEnd As Long
    Static DeltaAngle As Single, DeltaAr As Single
    Static Alpha As Single, Beta As Single
    Static TR As Single, TI As Single, AR As Single, AI As Single
    
    For I = 0 To (NumSamples - 1)
        j = ReversedBits(I) 'I saved time here by pre-calculating all these values
        RealOut(j) = RealIn(I)
        ImagOut(j) = 0 'Since this array is static, gotta make sure it's clear
    Next
    
    BlockEnd = 1
    BlockSize = 2
    
    Do While BlockSize <= NumSamples
        DeltaAngle = AngleNumerator / BlockSize
        Alpha = Sin(0.5 * DeltaAngle)
        Alpha = 2! * Alpha * Alpha
        Beta = Sin(DeltaAngle)
        
        I = 0
        Do While I < NumSamples
            AR = 1!
            AI = 0!
            
            j = I
            For n = 0 To BlockEnd - 1
                k = j + BlockEnd
                TR = AR * RealOut(k) - AI * ImagOut(k)
                TI = AI * RealOut(k) + AR * ImagOut(k)
                RealOut(k) = RealOut(j) - TR
                ImagOut(k) = ImagOut(j) - TI
                RealOut(j) = RealOut(j) + TR
                ImagOut(j) = ImagOut(j) + TI
                DeltaAr = Alpha * AR + Beta * AI
                AI = AI - (Alpha * AI - Beta * AR)
                AR = AR - DeltaAr
                j = j + 1
            Next
            
            I = I + BlockSize
        Loop
        
        BlockEnd = BlockSize
        BlockSize = BlockSize * 2
    Loop
End Sub


