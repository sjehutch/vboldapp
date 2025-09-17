// =DOC FILE Interface for the DMaxShootingRate class
// =DOC TEXT Declares DMaxShootingRate class to determine the maximum allowed shooting rate for introducing ProFume through a specific shooting line
//////////////////////////////////////////////////////////////////

// MaxShootingRate.h: interface for the DMaxShootingRate class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_MAXSHOOTINGRATE_H__CE204134_8797_4296_9A60_D0FD0C88551E__INCLUDED_)
#define AFX_MAXSHOOTINGRATE_H__CE204134_8797_4296_9A60_D0FD0C88551E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


#include <afxtempl.h>
#include "ValidateShootingLine.h"


// =DOC SECTION Struct tagRelHumidityIndex ( REL_HUMIDITY_DATA, *LPREL_HUMIDITY_DATA )
// =DOC SECTION Carrier structure containing relative humidity information
// =DOC TEXT struct tagRelHumidityIndex extends nothing
// =DOC TEXT Carrier structure containing relative humidity information
//////////////////////////////////////////////////////////////////
typedef struct tagRelHumidityIndex
{
   // =DOC SECTION Member nRelativeHumidity: int
   // =DOC SECTION Relative humidity value
   // =DOC TEXT nRelativeHumidity: int
   // =DOC TEXT Relative humidity value
   //////////////////////////////////////////////////////////////////
   int nRelativeHumidity;

   // =DOC SECTION Member dIndex: double
   // =DOC SECTION Relative humidity index value
   // =DOC TEXT dIndex: double
   // =DOC TEXT Relative humidity index value
   //////////////////////////////////////////////////////////////////
   double dIndex;

   // =DOC SECTION struct constructor
   // =DOC SECTION Initializes the structure
   // =DOC TEXT struct constructor extends nothing
   // =DOC TEXT Initializes the structure.
   //////////////////////////////////////////////////////////////////
   tagRelHumidityIndex()
   {
      nRelativeHumidity = 0;
      dIndex = 0.0;
   };

} REL_HUMIDITY_DATA, *LPREL_HUMIDITY_DATA;


// Array to store relative humidity indexes
typedef CArray< REL_HUMIDITY_DATA, REL_HUMIDITY_DATA& > REL_HUMIDITY_DATA_ARRAY;


// =DOC SECTION Class DMaxShootingRate
// =DOC SECTION Determines the maximum allowed shooting rate for introducing ProFume through a specific shooting line
// =DOC TEXT class DMaxShootingRate extends nothing
// =DOC TEXT Determines the maximum allowed shooting rate for introducing ProFume through a specific shooting line
//////////////////////////////////////////////////////////////////
class DMaxShootingRate  
{
public:
	// =DOC SECTION Method GetMaxShootingRate( LPMAXSHOOTINGRATE_DATA lpdataMaxRate ): BOOL
	// =DOC SECTION Function calculates the maximum allowed shooting rate for a shooting line, given information about the environment.
	// =DOC TEXT public GetMaxShootingRate( LPMAXSHOOTINGRATE_DATA lpdataMaxRate ): BOOL
	// =DOC TEXT LPMAXSHOOTINGRATE_DATA lpdataMaxRate: Transfer structure
   // =DOC TEXT Returns BOOL: TRUE/FALSE
	// =DOC TEXT Function calculates the maximum allowed shooting rate for a shooting line, given information about the environment.
	//////////////////////////////////////////////////////////////////
	BOOL GetMaxShootingRate( LPMAXSHOOTINGRATE_DATA lpdataMaxRate );

   // =DOC SECTION Class constructor
   // =DOC SECTION Initializes the class
   // =DOC TEXT class constructor extends nothing
   // =DOC TEXT Initializes the class.
   //////////////////////////////////////////////////////////////////
	DMaxShootingRate();

   // =DOC SECTION Class destructor
   // =DOC SECTION Cleans up the class.
   // =DOC TEXT class destructor extends nothing
   // =DOC TEXT Cleans up the class.
   //////////////////////////////////////////////////////////////////
	virtual ~DMaxShootingRate();

private:
	// =DOC SECTION Method GetMaxShootingRate(const double &dFanCapacity, const int &dRelativeHumidity): double
	// =DOC SECTION Function calculates the maximum allowed shooting rate for a shooting line, given information about the line and the environment.
	// =DOC TEXT public GetMaxShootingRate(const double &dFanCapacity, const double &dRelativeHumidityIndex): double
	// =DOC TEXT double &dFanCapacity: The fan capacity, in cubic meters per minute, associated with the shooting line.
	// =DOC TEXT double &dRelativeHumidityIndex: The relative humidity index of the area where the shooting line has been set up.
   // =DOC TEXT Returns double: The maximum allowed shooting rate for the shooting line.
	// =DOC TEXT Function calculates the maximum allowed shooting rate for a shooting line, given information about the line and the environment.
	//////////////////////////////////////////////////////////////////
	double GetMaxShootingRate( const double &dFanCapacity, 
                              const double &dRelativeHumidityIndex );

	// =DOC SECTION Method GetRelativeHumidityIndex(const int &nRelativeHumidity): double
	// =DOC SECTION Function retrieves the relative humidity index for the relative humidity input.
	// =DOC TEXT public GetRelativeHumidityIndex(const int &nRelativeHumidity): double
	// =DOC TEXT double &dRelativeHumidity: The relative humidity of the area where the shooting line has been set up.
   // =DOC TEXT Returns double: The relative humidity index value.
	// =DOC TEXT Function retrieves the relative humidity index for the relative humidity input.
	//////////////////////////////////////////////////////////////////
	double GetRelativeHumidityIndex( const int &nRelativeHumidity );

	// =DOC SECTION Method PopulateRelativeHumidityIndex(): void
	// =DOC SECTION Populate m_adataRelHumidity with relative humidity data.
	// =DOC TEXT public PopulateRelativeHumidityIndex(): void
	// =DOC TEXT Returns void: No return value
	// =DOC TEXT Function populates the relative humidity data array.
	//////////////////////////////////////////////////////////////////
	void PopulateRelativeHumidityIndex();

   // =DOC SECTION Member m_adataRelHumidity: REL_HUMIDITY_DATA_ARRAY
   // =DOC SECTION Array of relative humidity data
   // =DOC TEXT m_adataRelHumidity: REL_HUMIDITY_DATA_ARRAY
   // =DOC TEXT Array of relative humidity data populated by the PopulateRelativeHumidityIndex() function.
   //////////////////////////////////////////////////////////////////
   REL_HUMIDITY_DATA_ARRAY m_adataRelHumidity;
};

#endif // !defined(AFX_MAXSHOOTINGRATE_H__CE204134_8797_4296_9A60_D0FD0C88551E__INCLUDED_)
