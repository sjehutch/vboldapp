// GlobalFunctions.cpp: implementation of the generic global functions in the dpcalc project
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "GlobalEnums.h"
#include "GlobalFunctions.h"
#include <math.h>


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif



double GetFinalDose( const double &dIntercept, 
                     const double &dSlope, 
                     const double &dProbit, 
                     const double &dSafetyFactor )
{
   // Equation for Target Dosage, per Part I Step 4, on 2002-07-10
   //
   //                    probit - intercept
   //  n               [ ------------------ ]
   // C TTargetDose = e        slope

   double dDose = -1.0;

   if( dSlope == 0.0 )
   {
      // Can't divide by 0
      return -1;
   }

   dDose = exp( ( dProbit - dIntercept ) / dSlope );
   

   // Equation for Target Dosage, per Part I Step 5, on 2002-07-10
   //
   //  n          n
   // C TFinal = C TTargetDose * ConfidenceFactor
   
   return dDose * dSafetyFactor;
}


double GetInitialConcentration( const double &dHltEstimated, 
                                const double &dTimeExposure, 
                                const double &dDoseFinal, 
                                const double &dNValue )
{
   // Equation for Initial Target Concentration, per Part I Step 7, on 2002-07-10
   //
   //   n           FinalDose
   // Co  = ----------------------------
   //       HLTe           (-t / HLTe)
   //       ----- * [ 1 - 2            ]
   //       ln(2)
   //
   //
   //          n  1/n
   // Co = [ Co  ]

   double dInitMaxConc = -1.0;

   double d2RaisedPower = pow( 2, - dTimeExposure / dHltEstimated );

   if( dHltEstimated == 0 || d2RaisedPower == 1 )
   {
      dInitMaxConc = -1.0;
   }
   else
   {
      double dInitConc = dDoseFinal / ( ( ( dHltEstimated / log( 2 ) ) * ( 1 - d2RaisedPower ) ) );

      dInitMaxConc = pow( dInitConc, 1 / dNValue );
   }

   return dInitMaxConc;
}

//12/22/04-YYT-Added
//formula computes for initial concentration if user has defined a target CT.
double GetInitialConcentration2(const double &dHltOriginal,
                                const double &dTimeExposure, 
                                const double &UserTargetCT, 
                                const double &dNValue)
{
   // Equation for Initial Target Concentration, per Part I Step 7, on 2004-12-22
   // When there is a user-defined concentration
   //
   //				UserTargetCT
   // CoTarget  = ----------------------------
   //				HLTo           (-t / HLTo)
   //				----- * [ 1 - 2            ]
   //				ln(2)
   //
   //
   //          n  1/n
   // Co = [ Co  ]

   double dInitMaxConc = -1.0;

   double d2RaisedPower = pow( 2, - dTimeExposure / dHltOriginal );

   if( dHltOriginal == 0 || d2RaisedPower == 1 )
   {
      dInitMaxConc = -1.0;
   }
   else
   {
      double dInitConc = UserTargetCT / ( ( ( dHltOriginal / log( 2 ) ) * ( 1 - d2RaisedPower ) ) );

      dInitMaxConc = pow( dInitConc, 1 / dNValue );
   }

   return dInitMaxConc;
}


double GetEHLT(double dOHLT, double dNValue)
{
   // Equation for Initial Target Concentration, per Part I Step 6, on 2002-07-10
   //
   // HLTe = HLTo / n

   if( dNValue == 0.0 )
   {
      // Can't divide by 0
      return -1;
   }

   return dOHLT / dNValue;
}
