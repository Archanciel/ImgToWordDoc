import argparse

parser = argparse.ArgumentParser()
parser.add_argument("val", type=int, help="display a square of a given number")
parser.add_argument("power", type=int, nargs="?", default=2, help="power to apply to val. 2 if not specified.")
parser.add_argument("-v", "--verbosity", type=int, choices=[0, 1, 2], help="increase output verbosity")
args = parser.parse_args()
answer = args.val**args.power

if args.verbosity == 2:
   print("{} power {} equals {}".format(args.val, args.power, answer))
elif args.verbosity == 1:
   print("{}^{} == {}".format(args.val, args.power, answer))
else:
   print(answer)