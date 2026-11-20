import React, { useState } from 'react'
import { handleError } from '../utils';

function ExpenseForm({ addTransaction }) {

    const [expenseInfo, setExpenseInfo] = useState({
        amount: '',
        text: ''
    })

    const handleChange = (e) => {
        const { name, value } = e.target;
        const copyExpenseInfo = { ...expenseInfo };
        copyExpenseInfo[name] = value;
        setExpenseInfo(copyExpenseInfo);
    }

    const addExpenses = (e) => {
        e.preventDefault();
        const { amount, text } = expenseInfo;
        if (!amount || !text) {
            handleError('Please add Expense Details');
            return;
        }
        addTransaction(expenseInfo);
        setExpenseInfo({ amount: '', text: '' })
    }

    return (
        <div className='container expense-form-card'>
            <div className='form-card-heading'>
                <span className='form-icon'>+</span>
                <div><p className='card-label'>NEW ENTRY</p><h2>Add a transaction</h2></div>
            </div>
            <form onSubmit={addExpenses}>
                <div>
                    <label htmlFor='text'>What was it for?</label>
                    <input
                        onChange={handleChange}
                        type='text'
                        name='text'
                        placeholder='Enter your Expense Detail...'
                        value={expenseInfo.text}
                    />
                </div>
                <div>
                    <label htmlFor='amount'>Amount</label>
                    <input
                        onChange={handleChange}
                        type='number'
                        name='amount'
                        placeholder='Enter your Amount...'
                        value={expenseInfo.amount}
                    />
                </div>
                <button className='primary-button' type='submit'>Add transaction <span>→</span></button>
            </form>
        </div>
    )
}

export default ExpenseForm