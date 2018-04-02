#!/usr/bin/env python

import sys
import openpyxl 
import datetime

##
## replace multiple spaces in string with a single one.
def single_space( s ):
	return ' '.join(s.split())

##
## return a cell value as a string ...
def get_cell_value( cell ):

	v = cell.value
	if not v:
		return ''

	if type( v ) == str:
		return single_space( v )

	if type( v ) == float:
		rv = "%10.2f" % v
		return rv.strip()

	if type( v ) == long:
		return "%d" % v

	if type(v) == unicode:
		return single_space( v.encode('utf-8') )

	if type( v ) == datetime.datetime:
		return v.strftime( "%Y.%m.%d_%H:%M:%S" )

	return single_space( v )

##
## return a boolean true if row is <empty> ...
def empty_row( row ):
    for v in row:
        if v.value is not None:
            return False
    return True

##
## output with some control ...
def out( s ):
    if s:
        sys.stdout.write( s )

##
## recode a list of cells to a list of values ...
def recode( row ):
	rv=[]
	for x in row:
		c_val = get_cell_value( x )
		if c_val:
			rv.append( c_val )
		else: 
			rv.append( '' )
	return rv

def create_map( style, hdr0, hdr1 ):
	pl = [] 
	dim = [] 
	wgt = [] 
	cube = None
	for pt in price_tags:
		if pt in hdr0:
			pl.append( hdr0.index( pt ) )

	dim = [i for i, elem in enumerate(hdr0) if 'DIMENSIONS' in elem]
	grades = range( pl[0], dim[0] )

	for w in weight_tags:
		if w in hdr1:
			wgt.append( hdr1.index( w ) )
		
	rv = { "style": style, "price_lists": pl, "dimension": dim, "weight": wgt, "grades": grades, "hdr1": hdr1 }
	out( "%s\n" % rv )
	return rv

def walk_map( map, row ):
	for g in map['grades']:
		if row[g]:
			out( "%s," % map['style'] )
			out( "%s," % row[0] )
			out( "%s," % map['hdr1'][g] )
			out( "%s," % row[g] )
			out( "\n" )

price_tags = [
	"MSRP",
	"CDN PRICING",
	"CDN FABRIC PRICING",
	"CDN LEATHER PRICING",
	"MSRP FABRIC PRICING",
	"MSRP LEATHER PRICING"
]

weight_tags = [
	"(KG)",
	"(LBS)"
]

stop_tags = [
	"DIMENSIONS",
]

##
## main logic ...  
def main( args ):

	fn = args[1]
	wb = openpyxl.load_workbook( fn, data_only = True, read_only = True )

	for sn in wb.sheetnames:
##		print "::::::", fn, sn, wb[sn].max_column, wb[sn].max_row
		nhdrs = 0
		last_state = None
		state = "Searching"
		for r in wb[sn].rows:

			if empty_row( r ):
				continue

			v = recode( r )
			s = ",".join( v )
			if "STYLE: " in v[0]:
				state = "Found Style"
				style = v[0].strip( 'STYLE: ' )
##				print state, v[0]
				continue

			if "DESCRIPTION" == v[0]:
				state = "Found Description"
				nhdrs += 1
##				print state, nhdrs
				h0 = v
				continue

			if state == "Found Description":
				state = "Found Header"
				map = create_map( style, h0, v )
				state = "Expect Data"
				continue

			if state == "Expect Data":
				walk_map( map, v )
		else:
			last_state = state
			if not last_state == "Found Style":
				state = "Searching"

## a magic incantation ...
if __name__ == '__main__':
    main( sys.argv )
