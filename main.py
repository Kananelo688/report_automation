#!/usr/bin/env python
# coding: utf-8

#

import packet_loss as pl
import etl_core as ec
import argparse
from datetime import datetime, date




def main():
	parser = argparse.ArgumentParser(description='Report Automation Script.')
	parser.add_argument('--task',type=int,help='Specifies the kind of report to compile. 1 represents packet loss report,\
		2 represents ETL CORE of the Qlik Report, and 3 represents ETL RAN of the Qlik Report.')
	parser.add_argument('--report_filename',type=str,help='Must be folowed by the Report filename(with "xlsx" extension)')
	parser.add_argument('--verbose', action='store_true', help='Enable verbose output.')

	parser.add_argument('--end_date',type = str, help = 'Specifies the End date  on the Qlik Report(YYYY-MM-DD)')

	args = parser.parse_args()
	
	if args.task == 1:
		if args.verbose:
			print('\nReading Command-Line Arguments...',end = '' , flush= True)
	
		if args.report_filename is None and args.verbose:
			print('done.\nReport Filename not specified, Results will be saved to "Packetloss.xlsx"',flush=True)
			pl.compile_report(verbose=args.verbose)
		elif args.report_filename is not None:
			if args.verbose:
				print('done',flush = True)
			pl.compile_report(args.report_filename,verbose=args.verbose)
	elif args.task == 2:
		if args.verbose:
			print('Computing ETL Core Report:\nReading Command-line arguments...',end = ' ', flush= True)
		
		if args.end_date is None:
			raise parser.error('End date of the Qlik Spreadsheet where to insert new data must be specified as: "YYYY-MM-DD".')
		
		if args.verbose:
			print('done',flush=True)
		end_date = datetime.strptime(args.end_date, '%Y-%m-%d').date()
		ec.etl_core(end_date,
			verbose = args.verbose)

	elif args.task == 3:
		pass
	else:
		raise parser.error(f'Task must be 1,2 or 3. Task of {args.task} is not supported')

if __name__ == '__main__':
	main()