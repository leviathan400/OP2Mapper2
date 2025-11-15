$$rem Default template for C++ based DLLs
$$begin
void SetupUnits()
{
	// Create OP2Mapper generated units
	Unit x;
$$unit
	TethysGame::CreateUnit(x, $mapid, LOCATION($x+31, $y-1), $playerid, $weaponmapid, 0);
$$end
}
