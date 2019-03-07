using JUST.Shared.Classes;
using JUST.Shared.Utilities;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JUST.Shared.Tests
{
    [TestFixture()]
    public class UtilitiesTests
    {
        List<Employee> EmployeeList = new List<Employee>()
            {
                new Employee("1", "Enzo", "enzo@ferrari.com"),
                new Employee("2", "Dino", "dino@ferrari.com"),
                new Employee("3", "Carroll Shelby", "carroll@shelby.com"),
            };


        [SetUp]
        public void TestSetup()
        {
        }

        #region GetEmployeeInformation
        [TestCase]
        public void Program_GetEmployeeInformation_EmptyListReturnsEmptyEmployee()
        {
            var emptyEmployeeList = new List<Employee>();

            var result = EmployeeLookup.FindEmployeeFromAllEmployees(emptyEmployeeList, "Joe");
            Assert.IsInstanceOf(typeof(Employee), result);
            Assert.AreEqual(result.Name, string.Empty);
            Assert.AreEqual(result.EmailAddress, string.Empty);
            Assert.AreEqual(result.EmployeeId, string.Empty);
            Assert.AreEqual(result.Name, string.Empty);
        }

        [TestCase]
        public void Program_GetEmployeeInformation_EmployeeNotFoundInListReturnsEmptyEmployee()
        {
            var result = EmployeeLookup.FindEmployeeFromAllEmployees(EmployeeList, "Joe");
            Assert.IsInstanceOf(typeof(Employee), result);
            Assert.AreEqual(result.Name, string.Empty);
            Assert.AreEqual(result.EmailAddress, string.Empty);
            Assert.AreEqual(result.EmployeeId, string.Empty);
            Assert.AreEqual(result.Name, string.Empty);
        }

        [TestCase(0)]
        [TestCase(1)]
        [TestCase(2)]
        public void Program_GetEmployeeInformation_EmployeeFoundByEmployeeIdReturnsEmployee(int index)
        {
            var result = EmployeeLookup.FindEmployeeFromAllEmployees(EmployeeList, EmployeeList[index].EmployeeId);

            Assert.IsInstanceOf(typeof(Employee), result);
            Assert.AreEqual(result.Name, EmployeeList[index].Name);
            Assert.AreEqual(result.EmailAddress, EmployeeList[index].EmailAddress);
            Assert.AreEqual(result.EmployeeId, EmployeeList[index].EmployeeId);
            Assert.AreEqual(result.Name, EmployeeList[index].Name);
        }

        [TestCase(0)]
        [TestCase(1)]
        [TestCase(2)]
        public void Program_GetEmployeeInformation_EmployeeFoundByEmployeeNameReturnsEmployee(int index)
        {
            var result = EmployeeLookup.FindEmployeeFromAllEmployees(EmployeeList, EmployeeList[index].Name);

            Assert.IsInstanceOf(typeof(Employee), result);
            Assert.AreEqual(result.Name, EmployeeList[index].Name);
            Assert.AreEqual(result.EmailAddress, EmployeeList[index].EmailAddress);
            Assert.AreEqual(result.EmployeeId, EmployeeList[index].EmployeeId);
            Assert.AreEqual(result.Name, EmployeeList[index].Name);
        }

        [TestCase(0)]
        [TestCase(1)]
        [TestCase(2)]
        public void Program_GetEmployeeInformation_EmployeeFoundByLowerCaseEmployeeNameReturnsEmployee(int index)
        {
            var result = EmployeeLookup.FindEmployeeFromAllEmployees(EmployeeList, EmployeeList[index].Name.ToLower());

            Assert.IsInstanceOf(typeof(Employee), result);
            Assert.AreEqual(result.Name, EmployeeList[index].Name);
            Assert.AreEqual(result.EmailAddress, EmployeeList[index].EmailAddress);
            Assert.AreEqual(result.EmployeeId, EmployeeList[index].EmployeeId);
            Assert.AreEqual(result.Name, EmployeeList[index].Name);
        }

        [TestCase(0)]
        [TestCase(1)]
        [TestCase(2)]
        public void Program_GetEmployeeInformation_EmployeeFoundByUpperCaseEmployeeNameReturnsEmployee(int index)
        {
            var result = EmployeeLookup.FindEmployeeFromAllEmployees(EmployeeList, EmployeeList[index].Name.ToUpper());

            Assert.IsInstanceOf(typeof(Employee), result);
            Assert.AreEqual(result.Name, EmployeeList[index].Name);
            Assert.AreEqual(result.EmailAddress, EmployeeList[index].EmailAddress);
            Assert.AreEqual(result.EmployeeId, EmployeeList[index].EmployeeId);
            Assert.AreEqual(result.Name, EmployeeList[index].Name);
        }
        #endregion

    }
}
