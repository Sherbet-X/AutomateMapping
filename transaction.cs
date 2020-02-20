using System;
using System.Data;
using Oracle.DataAccess.Client; 
 
class SaveSample
{
  static void Main()
  {
    // Drop & Create MyTable as indicated in Database Setup, at beginning
    
    // This sample starts a transaction and creates a savepoint after
    // inserting one record into MyTable.
    // After inserting the second record it rollsback to the savepoint
    // and commits the transaction. Only the first record will be inserted
 
    string constr = "User Id=scott;Password=tiger;Data Source=oracle";
    OracleConnection con = new OracleConnection(constr);
    con.Open();
 
    OracleCommand cmd = con.CreateCommand();
 
    // Check the number of rows in MyTable before transaction
    cmd.CommandText = "SELECT COUNT(*) FROM MyTable";    
    int myTableCount = int.Parse(cmd.ExecuteScalar().ToString());
 
    // Print the number of rows in MyTable
    Console.WriteLine("myTableCount = " + myTableCount);
 
    // Start a transaction
    OracleTransaction txn = con.BeginTransaction(
      IsolationLevel.ReadCommitted);
 
    // Insert a row into MyTable
    cmd.CommandText = "INSERT INTO MyTable VALUES (1)";
    cmd.ExecuteNonQuery();
 
    // Create a savepoint
    txn.Save("MySavePoint");
 
    // Insert another row into MyTable
    cmd.CommandText = "insert into mytable values (2)";
    cmd.ExecuteNonQuery();
 
    // Rollback to the savepoint
    txn.Rollback("MySavePoint");
 
    // Commit the transaction
    txn.Commit();
 
    // Check the number of rows in MyTable after transaction
    cmd.CommandText = "SELECT COUNT(*) FROM MyTable";
    myTableCount = int.Parse(cmd.ExecuteScalar().ToString());
 
    // Prints the number of rows, should have increased by 1
    Console.WriteLine("myTableCount = " + myTableCount);
 
    txn.Dispose();
    cmd.Dispose();
 
    con.Close();
    con.Dispose();
  }
}