# OP2Mapper2

![Screenshot](https://wiki.outpost2.net/lib/exe/fetch.php?cache=&w=500&h=426&tok=19109a&media=outpost_2:helper_programs:mapper_screenshot.png)

Created by BlackBox and Hooman

https://wiki.outpost2.net/doku.php?id=outpost_2:helper_programs:mapper

Originally released 21/04/2005 (2.0 Alpha 1)

https://forum.outpost2.net/index.php/topic,1595.0.html

Visual Basic 6.0 project


Source code is from OP2Mapper2-source.rar (Modified date 13/09/2006). This was marked as 2.2.4 in the project file.


The backend code uses OP2Editor: https://github.com/OutpostUniverse/OP2Editor



# File Formats

OP2Mapper2 has three special file formats, .ctl, .tpl and .dat

- **.ctl** - Object definitions
- **.tpl** - Code generation template
- **.dat** - Contains unit placem and object placements



# OP2Mapper2 – CTL File Format Specification

This document explains the file format and purpose of the **CTL** files used by **OP2Mapper2**.

---

## Overview

OP2Mapper2 uses several `.ctl` data files to understand:

- units.ctl - Unit types (structures, vehicles, turrets)
- weapons.ctl - Weapon types and turret compatibility
- objects.ctl - Gaia objects (beacons, fumaroles, walls, tubes, etc.)
- terrains.ctl - Terrain type definitions and tile index ranges inside Outpost 2 tilesets

These files allow the mapper to interpret `.map` and `.dat` contents in a way that matches Outpost 2’s internal game logic.

---

# General CTL Conventions

### **Comments**
Lines beginning with `;` are ignored.

```
; This is a comment
```

### **INCLUDE**
CTL files may include other CTL files:

```
INCLUDE "filename.ctl"
```

- Filename is in quotes  
- Path is relative to the including file  
- Inclusion is recursive  

### **Whitespace Behavior**
- Leading/trailing spaces trimmed  
- Tabs removed  
- Empty lines ignored  

### **Fields**
Fields are comma-separated, strings in quotes:

```
28,"Command Center",3,2,2,0
```

The meaning of fields depends on the CTL type.

---

## `units.ctl` — Standard Outpost 2 Unit Definitions

Defines all **non-Gaia** units:

- Structures (CC, Agridome, Res, Storage, Mines, etc.)
- Vehicles (Cargo Truck, Convec, Miner, Scouts, Combat Vehicles)
- Turrets as units (for mapper placement)

These match `.map.dat` entries where `IsGaia = False`.

### **Format**
Depending on unit type, the number of fields varies:

```
MapID,"Name",ArtID
MapID,"Name",Width,Height,ArtID
MapID,"Name",Width,Height,VerTubeLoc,HorTubeLoc,ArtID
```

### **Field meanings**

| Field          | Meaning                                                   |
|----------------|-----------------------------------------------------------|
| **MapID**      | Outpost 2 internal object ID                             |
| **Name**       | Display name                                              |
| **Width/Height** | Footprint in tiles (buildings only)                    |
| **VerTubeLoc** | Tube attach point (vertical)                              |
| **HorTubeLoc** | Tube attach point (horizontal)                            |
| **ArtID**      | Tile index used by mapper                                 |

### Notes
- `units.ctl` does **not** specify turret compatibility.  
- That is handled separately in `weapons.ctl` via `HASTURRET`.

---

## `weapons.ctl` — Weapon Definitions & Turret Compatibility

Has two roles:

### 1. **Weapon definition lines**
```
MapID,"WeaponName",ArtID
```

| Field         | Meaning                           |
|---------------|-----------------------------------|
| **MapID**     | Weapon ID used in `.map.dat`      |
| **WeaponName**| Display name                      |
| **ArtID**     | Tile index for mapper             |

### 2. **Turret compatibility lines**
```
HASTURRET <UnitMapID>
```

This sets:

```
UnitDefs(UnitMapID).CanHaveTurret = True
```

---

## `objects.ctl` — Gaia Objects / Special Map Features

Defines all **Gaia (neutral) objects**, such as:

- Common/Rare ore beacons (all yield variants)
- Walls (Normal Wall, Microbe Wall, Magma Wall)
- Tubes (Gaia version)
- Fumaroles, Magma vents
- Rubble, wreckage, tech artifacts
- Other map-placed special tiles

These are used when `.map.dat` entries have `IsGaia = True`.

### **Format**
Several variants are permitted:

```
MapID,"Name",Type,ArtID
MapID,"Name",Type,ArtID,Extra1
MapID,"Name",Type,ArtID,Extra1,Extra2
MapID,"Name",Type,ArtID,Extra1,Extra2,Extra3
```

### **Field meanings**

| Field     | Explanation                                                   |
|-----------|---------------------------------------------------------------|
| **MapID** | Gaia base ID (81 = beacon, 18 = normal wall, etc.)           |
| **Name**  | Display name                                                  |
| **Type**  | Category flag (varies by object)                              |
| **ArtID** | Tile index                                                    |
| **Extra1–3** | Object parameters — interpretation depends on object type |

### **Example: Beacons**

All share MapID **81**:

```
81,"Common Beacon 1",0,1749,0,2,-1
81,"Common Beacon 2",0,1750,0,1,-1
...
81,"Rare Beacon 3",0,1754,1,0,-1
```

Meaning:

| Extra | Meaning                      |
|-------|-------------------------------|
| Extra1 | 0=Common, 1=Rare              |
| Extra2 | Yield (0 low, 1 med, 2 high)  |
| Extra3 | Variant index                 |

Gaia resolution is done via:

```
MapID + Extra1 + Extra2 + Extra3
```

---

## `terrains.ctl` — Terrain Types & Tile Ranges

This file defines **how tile indices (0–2011)** map to **terrain types** for a given tileset.

Outpost 2 tilesets (e.g., `well00`) have:

- 2012 tiles total  
- Multiple terrain types (commonly 4: sand, rock, lava, etc.)  
- Special tiles (tubes, walls, cracks, etc.)

### **Example Concept**

```
NumTerrains=4
```

Meaning: 4 terrain types exist.

Each terrain block specifies:

```
[Terrain0]
StartTile=0
EndTile=437
```

This means tile indices **0–437** belong to terrain type 0.

Additional entries define where special tiles appear (wall tiles, tube tiles, etc.).

### **Purpose of terrains.ctl**

It allows the mapper to:

- Know which tile range corresponds to each terrain type  
- Identify special tile indices (walls/tubes) *per terrain type*  
- Select correct tile graphics when placing walls/tubes on different terrain types  

### **Example**

Tile index **331** = first Normal Wall tile on terrain type 0 (sand).

Mapper2 uses this CTL to keep the correct tile variant for each terrain type.

---

# `.dat` Interaction Summary

When loading `.dat`:

### **Non-Gaia units (`IsGaia=False`)**
Resolved via:
```
units.ctl ? UnitDefs[MapID]
```

### **Gaia objects (`IsGaia=True`)**
Resolved via:
```
objects.ctl ? Objects[MapID + Extras]
```

### **Weapons**
Resolved via:
```
weapons.ctl ? Weapons[WeaponMapID]
```
Plus turret flags applied to UnitDefs.

### **Terrain tiles**
Determined via:
```
terrains.ctl ? terrain ranges and special tile indices
```

---

# Summary Table

| CTL file        | Purpose                                       | Used for resolving                     |
|-----------------|-----------------------------------------------|-----------------------------------------|
| `units.ctl`     | Structures / vehicles                         | `MapID`                                 |
| `weapons.ctl`   | Weapon names + turret compatibility           | `MapID`                                 |
| `objects.ctl`   | Gaia objects (beacons, walls, tubes, etc.)    | `MapID + Extra1/2/3`                    |
| `terrains.ctl`  | Terrain types & tile index ranges             | Tile index + terrain type               |

These files collectively describe nearly all Outpost-2 related data needed to correctly interpret and edit a map.



# OP2Mapper2 – DAT File Format Specification

The binary `.dat` file format produced by **OP2Mapper2** and used alongside Outpost 2 `.map` files. The `.dat` file contains **unit placements**, **Gaia object placements**, and extra metadata for OP2Mapper2 that is not stored in the `.map` tiles and cell type file. Is is generated when using the 'Place Object' function.

---

# Overview

Outpost 2 uses two files for a map:

| File | Purpose |
|------|---------|
| `*.map` | Tilemap (terrain tiles, cell types of tiles) |
| `*.dat` | Unit placements, Gaia objects, and attributes |

Mapper2 maintains its own `.dat` format, which is **not the same** as the Outpost 2 mission DLL format.  It is a lightweight OP2Mapper2-era binary structure.

---

# High-Level Structure

The `.dat` file contains:

```
[NumberOfUnitRecords : 4 bytes]
[Record 0]
[Record 1]
...
[Record N-1]
```

Where each record is a **fixed binary struct**, optionally extended for Gaia objects.

---

# Binary Record Format

Each record is:

| Field | Type | Size | Description |
|-------|------|------|-------------|
| `LocX` | Long (Int32) | 4 | X tile coordinate |
| `LocY` | Long (Int32) | 4 | Y tile coordinate |
| `PlayerNum` | Long (Int32) | 4 | Owner (0–6, or 0 for Gaia objects) |
| `UnitMapId` | Long (Int32) | 4 | ID from `units.ctl` **or** `objects.ctl` |
| `WeaponMapId` | Long (Int32) | 4 | ID from `weapons.ctl` (0 = none) |
| `IsGaia` | VB6 Boolean | 2 | 0 = false, -1 = true |

If `IsGaia = True`, three more **Integer (Int16)** values follow:

| Field | Type | Size | Description |
|--------|------|------|-------------|
| `Extra1` | Integer (Int16) | 2 | Extra attribute 1 (Gaia object subtype) |
| `Extra2` | Integer (Int16) | 2 | Extra attribute 2 (yield, variant, etc.) |
| `Extra3` | Integer (Int16) | 2 | Extra attribute 3 (depends on object) |

### Total Record Size

| Type | Size |
|------|-------|
| **Non-Gaia record** | 4+4+4+4+4 +2 = **22 bytes** |
| **Gaia record** | 22 + 2+2+2 = **28 bytes** |

The file begins with a **4-byte Int32 count**, so full size is:

```
4 + (RecordCount × 22 or 28)
```

---

# Detailed Field Semantics

## `LocX`, `LocY`

Tile coordinates on the Outpost 2 map (0–511 generally).

## `PlayerNum`

- 0–6 ? Human or AI player  
- 0 with `IsGaia=True` ? Neutral map-placed object  
- 0 with `IsGaia=False` ? Player 0 (normal unit for player 0)

## `UnitMapId`

### If `IsGaia=False`  
Resolved through `units.ctl`.

### If `IsGaia=True`  
Resolved through `objects.ctl` in combination with:

- `Extra1`
- `Extra2`
- `Extra3`

Example: All ore beacons use UnitMapId = 81.

## `WeaponMapId`

Weapon ID resolved through `weapons.ctl`.

- 0 means "None"
- Mapper2 will display turret names and verify turret compatibility

## `IsGaia`

VB6 Boolean values:

| Value | Meaning |
|--------|----------|
| `0` | Not Gaia |
| `-1` | Gaia object |

## `Extra1–Extra3` (Gaia only)

These fields are object-specific. Examples:

### Beacon objects
```
Extra1 = 0 or 1 (common/rare)
Extra2 = 0–2 (yield)
Extra3 = variant index
```

### Walls / tubes
Often used to distinguish terrain-specific variants.

---

# Example Binary Layout (hex)

A *single non-Gaia unit*:

```
13 00 00 00   LocX = 19
03 00 00 00   LocY = 3
00 00 00 00   Player 0
01 00 00 00   UnitMapId = 1  (Cargo Truck)
00 00 00 00   WeaponMapId = 0
00 00         IsGaia = False
```

A *Gaia beacon object*:

```
14 00 00 00   LocX = 20
05 00 00 00   LocY = 5
00 00 00 00   Player = 0 (but Gaia flag set)
51 00 00 00   UnitMapId = 81 (beacon base ID)
00 00 00 00   Weapon = none
FF FF         IsGaia = -1 (True)
00 00         Extra1 = 0 (common)
02 00         Extra2 = 2 (high yield)
FF FF         Extra3 = -1
```

---

# Parsing Rules

### Step 1 — Read number of records

```
count = br.ReadInt32()
```

### Step 2 — Read each record

- Always read first 22 bytes
- Check the 6th field (`IsGaia`)
- If Gaia ? read 6 more bytes

### Step 3 — Lookup definitions

| Field | Lookup Source |
|--------|----------------|
| UnitMapId (Non-Gaia) | units.ctl |
| UnitMapId + Extras | objects.ctl |
| WeaponMapId | weapons.ctl |

---

# Relationship Between `.map` and `.map.dat`

| File | Contains |
|------|----------|
| `.map` | Terrain only (tile IDs, 512×512) |
| `.map.dat` | All units, Gaia objects, weapons, extra attributes |

Both are required for a functional Outpost 2 mission map.

---

# Summary

The OP2Mapper2 `.dat` format is:

- Simple
- Fixed-record
- VB6-aligned
- Purpose-built for map editing, not used by Outpost 2 directly

| Section | Purpose |
|---------|----------|
| Header | Number of records |
| Records | Units and Gaia objects |
| Extras | Gaia parameter fields |

