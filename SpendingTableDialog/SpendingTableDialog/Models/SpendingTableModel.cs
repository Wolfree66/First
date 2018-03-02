using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpendingTableDialog.Models
{
    public class SpendingTableModel
    {

        public SpendingTableModel(
            double fare, string fareMoneyTransferType,
            double dailyAllowance, string dailyAllowanceMoneyTransferType,
            double hotelExpenses, string hotelExpensesMoneyTransferType,
            double forceMajorExpenses, string forceMajorExpensesMoneyTransferType)
        {
            this.Fare = fare;
            this.FareMoneyTransferType = fareMoneyTransferType;
            this.DailyAllowance = dailyAllowance;
            this.DailyAllowanceMoneyTransferType = dailyAllowanceMoneyTransferType;
            this.HotelExpenses = hotelExpenses;
            this.HotelExpensesMoneyTransferType = hotelExpensesMoneyTransferType;
            this.Fare = fare;
            this.ForceMajorExpensesMoneyTransferType = forceMajorExpensesMoneyTransferType;

        }

        /// <summary>
        /// Транспортные расходы
        /// </summary>
        public double Fare { get; private set; }

        /// <summary>
        /// Способ получения денег на оплату транспортных расходов
        /// </summary>
        public string FareMoneyTransferType { get; private set; }

        /// <summary>
        /// Суточные
        /// </summary>
        public double DailyAllowance { get; private set; }

        /// <summary>
        /// Способ получения денег на оплату суточных
        /// </summary>
        public string DailyAllowanceMoneyTransferType { get; private set; }

        /// <summary>
        /// Расходы на проживание
        /// </summary>
        public double HotelExpenses { get; private set; }

        /// <summary>
        /// Способ получения денег на оплату расходов на проживание
        /// </summary>
        public string HotelExpensesMoneyTransferType { get; private set; }

        /// <summary>
        /// Непредвиденные расходы
        /// </summary>
        public double ForceMajorExpenses { get; private set; }

        /// <summary>
        /// Способ получения денег на оплату непредвиденных расходов
        /// </summary>
        public string ForceMajorExpensesMoneyTransferType { get; private set; }
    }
}
