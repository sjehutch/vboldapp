// =DOC FILE Interface for the DShootingRate class
// =DOC TEXT Declares DShootingRate class to determine the amount of profume required for a fumigation

// ShootingRate.h: interface for the DShootingRate class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DSHOOTINGRATE_H__4551F533_27F4_4D52_9761_5D75AC171AF5__INCLUDED_)
#define AFX_DSHOOTINGRATE_H__4551F533_27F4_4D52_9761_5D75AC171AF5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#include "ValidateShootingLine.h"


// =DOC SECTION Class DShootingRate
// =DOC SECTION Determine the shooting rate for a shooting line
// =DOC TEXT class DShootingRate extends nothing
// =DOC TEXT Determine the shooting rate for a shooting line
//////////////////////////////////////////////////////////////////
class DShootingRate  
{
public:
	
   // =DOC SECTION Method GetShootingRate( LPMAXSHOOTINGRATE_DATA lpdataShootingRate ): BOOL
   // =DOC SECTION Determine the shooting rate for a shooting line
	// =DOC TEXT public GetShootingRate(  LPMAXSHOOTINGRATE_DATA lpdataShootingRate  ): BOOL
	// =DOC TEXT LPMAXSHOOTINGRATE_DATA lpdataShootingRate: Transfer structure
	// =DOC TEXT Returns BOOL: TRUE (No Errors) / FALSE (Errors)
	// =DOC TEXT Determine the shooting rate for a shooting line
	//////////////////////////////////////////////////////////////////
	BOOL GetShootingRate( LPSHOOTINGRATE_DATA lpdataRate );

	// =DOC SECTION Method DShootingRate(): constructor
	// =DOC SECTION Class constructor
	// =DOC TEXT public DShootingRate()
	// =DOC TEXT Class constructor
	//////////////////////////////////////////////////////////////////
	DShootingRate();

	// =DOC SECTION Method ~DShootingRate(): destructor
	// =DOC SECTION Class destructor
	// =DOC TEXT public virtual ~DShootingRate()
	// =DOC TEXT Class destructor
	//////////////////////////////////////////////////////////////////
   virtual ~DShootingRate();

private:
	double GetTimeToEmptyCylinder(const double &dInitialFlowRate, 
		                          const double &dNumCylinders);
   // =DOC SECTION Method GetShootingRate( const double &dCylinderTemp, const double &dHoseLength, const double &dHoseDiameter ): double
   // =DOC SECTION Determine the shooting rate for a shooting line
	// =DOC TEXT public GetShootingRate( const double &dCylinderTemp, const double &dHoseLength, const double &dHoseDiameter ): double
	// =DOC TEXT dCylinderTemp: Cylinder temperature, C
	// =DOC TEXT dHoseLength: Hose length, M
	// =DOC TEXT dHoseDiameter: Hose diameter, mm
	// =DOC TEXT Returns double: shooting rate
	// =DOC TEXT Determine the shooting rate for a shooting line
	//////////////////////////////////////////////////////////////////
	double GetShootingRate( const double &dCylinderTemp, 
                           const double &dOneFourthLength,
                           const double &dOneEightLength );

};

#endif // !defined(AFX_DSHOOTINGRATE_H__4551F533_27F4_4D52_9761_5D75AC171AF5__INCLUDED_)
