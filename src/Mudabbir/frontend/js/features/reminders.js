/**
 * Mudabbir - Reminders Feature Module
 *
 * Created: 2026-02-05
 * Extracted from app.js as part of componentization refactor.
 *
 * Contains reminder-related state and methods:
 * - Reminder CRUD operations
 * - Reminder panel management
 * - Time formatting
 */

window.Mudabbir = window.Mudabbir || {};

window.Mudabbir.Reminders = {
    name: 'Reminders',
    /**
     * Get initial state for Reminders
     */
    getState() {
        return {
            showReminders: false,
            reminders: [],
            reminderInput: '',
            reminderLoading: false
        };
    },

    /**
     * Get methods for Reminders
     */
    getMethods() {
        return {
            /**
             * Handle reminders list
             */
            handleReminders(data) {
                this.reminders = data.reminders || [];
                this.reminderLoading = false;
            },

            /**
             * Handle reminder added
             */
            handleReminderAdded(data) {
                this.reminders.push(data.reminder);
                this.reminderInput = '';
                this.reminderLoading = false;
                this.showToast('Reminder set!', 'success');
            },

            /**
             * Handle reminder deleted
             */
            handleReminderDeleted(data) {
                this.reminders = this.reminders.filter(r => r.id !== data.id);
            },

            /**
             * Handle reminder triggered (notification)
             */
            handleReminderTriggered(data) {
                const reminder = data.reminder;
                this.showToast(`Reminder: ${reminder.text}`, 'info');
                this.addMessage('assistant', `Reminder: ${reminder.text}`);

                // Remove from local list
                this.reminders = this.reminders.filter(r => r.id !== reminder.id);

                // Try desktop notification
                if (Notification.permission === 'granted') {
                    new Notification('Mudabbir Reminder', {
                        body: reminder.text,
                        icon: "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'%3E%3Cdefs%3E%3ClinearGradient id='g' x1='4' y1='4' x2='28' y2='28' gradientUnits='userSpaceOnUse'%3E%3Cstop offset='0' stop-color='%237CF9FF'/%3E%3Cstop offset='0.55' stop-color='%2322D3EE'/%3E%3Cstop offset='1' stop-color='%231D4ED8'/%3E%3C/linearGradient%3E%3C/defs%3E%3Cg fill='none' stroke='url(%23g)' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M12.4 6.8c-3.7 0-6.8 2.9-6.8 6.5c0 .8.2 1.5.4 2.2c-1.5 1.2-2.4 3-2.4 5.1c0 3.4 2.7 6.2 6.1 6.2h3.4c1 1.2 2.6 2.1 4.4 2.1c1.8 0 3.4-.8 4.4-2.1h2.6c3.4 0 6.1-2.8 6.1-6.2c0-2.1-.9-3.9-2.4-5.1c.3-.7.4-1.4.4-2.2c0-3.6-3-6.5-6.8-6.5c-1.6 0-3 .5-4.2 1.3c-1.2-.8-2.6-1.3-4.2-1.3z'/%3E%3Cpath d='M17.5 8v18.9'/%3E%3C/g%3E%3C/svg%3E"
                    });
                }
            },

            /**
             * Open reminders panel
             */
            openReminders() {
                this.showReminders = true;
                this.reminderLoading = true;
                socket.send('get_reminders');

                // Request notification permission
                if (Notification.permission === 'default') {
                    Notification.requestPermission();
                }

                this.$nextTick(() => {
                    if (window.refreshIcons) window.refreshIcons();
                });
            },

            /**
             * Add a reminder
             */
            addReminder() {
                const text = this.reminderInput.trim();
                if (!text) return;

                this.reminderLoading = true;
                socket.send('add_reminder', { message: text });
                this.log(`Setting reminder: ${text}`, 'info');
            },

            /**
             * Delete a reminder
             */
            deleteReminder(id) {
                socket.send('delete_reminder', { id });
            },

            /**
             * Format reminder time for display
             */
            formatReminderTime(reminder) {
                const date = new Date(reminder.trigger_at);
                return date.toLocaleString(undefined, {
                    month: 'short',
                    day: 'numeric',
                    hour: '2-digit',
                    minute: '2-digit'
                });
            }
        };
    }
};

window.Mudabbir.Loader.register('Reminders', window.Mudabbir.Reminders);
