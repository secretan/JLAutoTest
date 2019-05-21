
# serial
import serial;
# thread
import threading;
# excel file
import xlrd;

# serial init
port = serial.Serial();
port.baudrate = 115200;
port.bytesize = serial.EIGHTBITS;
port.stopbits = serial.STOPBITS_ONE;
port.parity = serial.PARITY_NONE;
port.port = "COM21";

try:
    port.open();
except serial.SerialException as err:
    print(err.errno);
    exit();

def read_thread():


# thread init
#read_thread = threading.
threading.Thread(target=read_thread, name=rthread);

# excel file init
cmd = xlrd.open_workbook("cmd.xlsx");
# table
cmd_table = cmd.sheet_by_index(0);

def update_data_with_zero(dat):
    dlen = len(dat);
    for i in range(0, dlen):
        if (dat[i] == ""):
            dat[i] = "0";
    dat
    i

for i in range(0,cmd_table.nrows):
    cmd_pro = cmd_table.row_values(i);
    update_data_with_zero(cmd_pro);
    cmd_data = bytearray(list(map(int, cmd_pro)));
    port.write(cmd_data);
    print(cmd_data);



# serial deinit
port.close();
