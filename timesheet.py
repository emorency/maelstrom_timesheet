import gspread
import datetime
import argparse
import sys


def add_arguments(parser):
    """
    Add variable command specific options
    """
    parser.add_argument('project',
                        help='Project name: [bioshare, clsa, cpac, bbmri, ialsa, cihr, interconnect, mrc, p3g, maelstrom]')
    parser.add_argument('--coordination', '-c', required=False, help='Coordination and Meetings')
    parser.add_argument('--dev_other', '-dev', required=False, help='Software Development Other')
    parser.add_argument('--datashield', '-d', required=False, help='DataShield Software Development')
    parser.add_argument('--opal', '-op', required=False, help='Opal Software Development')
    parser.add_argument('--onyx', '-on', required=False, help='Onyx Software Development')
    parser.add_argument('--mica', '-m', required=False, help='Mica Software Development')
    parser.add_argument('--day', type=int, required=False, help='Day of the month (enter a number less than 0 to enter time for a previous day)')
    parser.add_argument('--month', required=False, help='Month of the year')
    # parser.add_argument('--interactive', '-i', required=False, help='Interactive mode')


def do_command(args):
    """
    Execute variable command
    """
    # Build and send request
    try:
        # Login with your Google account
        gc = gspread.login('@p3g.org', '')
        # Open a worksheet from spreadsheet with one shot
        #wks = gc.open("Maelstrom FT").sheet1
        sh = gc.open_by_key('0Ai3lwpfy7yHwdFBUNExYN3FoRGxjT1YwUXFrQ0JTX1E')

        wks = sh.get_worksheet(10)

        # if not args.interactive:

        now = datetime.datetime.now()
        day = now.day
        month = now.month
        year = now.year

        if args.day and args.day > 0:
            day = args.day
        elif args.day and args.day < 0:
            # day = day + args.day
            now = now + datetime.timedelta(days=args.day)
            day = now.day
            month = now.month
            year = now.year

        if args.month:
            month = args.month

        currdate = '%s/%s/%s' % (month, day, year)

        print("Update " + args.project + " for %s with (y/n) ?" % currdate)
        confirmed = sys.stdin.readline().rstrip().strip()
        if confirmed == "n":
            print("Aborted")
            sys.exit(2)

        cell = wks.find(currdate)

        print("Found something at R%sC%s" % (cell.row, cell.col))

        row = cell.row

        # PROJECT
        if args.project == "bioshare":
            col = 3
        elif args.project == "clsa":
            col = 5
        elif args.project == "cpac":
            col = 6
        elif args.project == "bbmri":
            col = 7
        elif args.project == "ialsa":
            col = 8
        elif args.project == "cihr":
            col = 9
        elif args.project == "interconnect":
            col = 10
        elif args.project == "mrc":
            col = 11
        elif args.project == "p3g":
            col = 12
        elif args.project == "maelstrom":
            col = 13

        # TASK
        if args.coordination:
            print("Coordination = " + args.coordination)
            wks.update_cell(row + 3, col, args.coordination)

        if args.dev_other:
            print("Dev Other = " + args.dev_other)
            wks.update_cell(row + 4, col, args.dev_other)

        if args.datashield:
            print("Datashield = " + args.datashield)
            wks.update_cell(row + 5, col, args.datashield)

        if args.opal:
            print("Opal = " + args.opal)
            wks.update_cell(row + 6, col, args.opal)

        if args.onyx:
            print("Onyx = " + args.onyx)
            wks.update_cell(row + 7, col, args.onyx)

        if args.mica:
            print("Mica = " + args.mica)
            wks.update_cell(row + 8, col, args.mica)

    except Exception, e:
        print e

def add_subcommand(name, help, add_args_func, default_func):
    """
    Make a sub-parser, add default arguments to it, add sub-command arguments and set the sub-command callback function.
    """
    subparser = subparsers.add_parser(name, help=help)
    # add_arguments(subparser)
    add_args_func(subparser)
    subparser.set_defaults(func=default_func)


# Parse arguments
parser = argparse.ArgumentParser(description='Timesheet command line.')
subparsers = parser.add_subparsers(title='sub-commands',
                                   help='Available sub-commands. Use --help option on the sub-command '
                                        'for more details.')

# Add subcommands
add_subcommand('today', 'Update Maelstrom Timesheet', add_arguments, do_command)

# Execute selected command
args = parser.parse_args()
args.func(args)
