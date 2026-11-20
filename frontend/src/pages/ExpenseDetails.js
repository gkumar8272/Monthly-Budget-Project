import React from 'react'

function ExpenseDetails({ incomeAmt, expenseAmt }) {
    const balance = incomeAmt - expenseAmt;

    return (
        <section className='summary-panel'>
            <div className='balance-block'>
                <p className='card-label'>AVAILABLE BALANCE</p>
                <strong>₹ {balance.toLocaleString('en-IN')}</strong>
                <span className='balance-note'>Across your recorded transactions</span>
            </div>
            <div className="amounts-container">
                <div className='amount-item'>
                    <span className='amount-dot income-dot'></span>
                    <div><span>Income</span><strong className="income-amount">₹{incomeAmt.toLocaleString('en-IN')}</strong></div>
                </div>
                <div className='amount-item'>
                    <span className='amount-dot expense-dot'></span>
                    <div><span>Expenses</span><strong className="expense-amount">₹{expenseAmt.toLocaleString('en-IN')}</strong></div>
                </div>
            </div>
        </section>
    )
}

export default ExpenseDetails