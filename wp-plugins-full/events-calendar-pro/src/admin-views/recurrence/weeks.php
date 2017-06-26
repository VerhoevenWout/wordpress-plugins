<div class="recurrence-row custom-recurrence-weeks" role="group" aria-label="<?php esc_html_e( 'Event Series Recurrence Day of Week', 'tribe-events-calendar-pro' ); ?>">
	<span class="tribe-field-inline-text first-label-in-line"><?php echo esc_html_x( 'On:', 'Begins the line indicating what day of the week the event will occur on', 'tribe-events-calendar-pro' ); ?></span>
	<label>
		<input type="checkbox" name="recurrence[<?php echo esc_attr( $rule_type ); ?>][][custom][week][day][]" data-field="custom-week-day" value="1" {{tribe_checked_if_in '1' custom.week.day}}/>
		<?php esc_html_e( 'M', 'tribe-events-calendar-pro' ); ?>
	</label>
	<label>
		<input type="checkbox" name="recurrence[<?php echo esc_attr( $rule_type ); ?>][][custom][week][day][]" data-field="custom-week-day" value="2" {{tribe_checked_if_in '2' custom.week.day}}/>
		<?php esc_html_e( 'Tu', 'tribe-events-calendar-pro' ); ?>
	</label>
	<label>
		<input type="checkbox" name="recurrence[<?php echo esc_attr( $rule_type ); ?>][][custom][week][day][]" data-field="custom-week-day" value="3" {{tribe_checked_if_in '3' custom.week.day}}/>
		<?php esc_html_e( 'W', 'tribe-events-calendar-pro' ); ?>
	</label>
	<label>
		<input type="checkbox" name="recurrence[<?php echo esc_attr( $rule_type ); ?>][][custom][week][day][]" data-field="custom-week-day" value="4" {{tribe_checked_if_in '4' custom.week.day}}/>
		<?php esc_html_e( 'Th', 'tribe-events-calendar-pro' ); ?>
	</label>
	<label>
		<input type="checkbox" name="recurrence[<?php echo esc_attr( $rule_type ); ?>][][custom][week][day][]" data-field="custom-week-day" value="5" {{tribe_checked_if_in '5' custom.week.day}}/>
		<?php esc_html_e( 'F', 'tribe-events-calendar-pro' ); ?>
	</label>
	<label>
		<input type="checkbox" name="recurrence[<?php echo esc_attr( $rule_type ); ?>][][custom][week][day][]" data-field="custom-week-day" value="6" {{tribe_checked_if_in '6' custom.week.day}}/>
		<?php esc_html_e( 'Sa', 'tribe-events-calendar-pro' ); ?>
	</label>
	<label>
		<input type="checkbox" name="recurrence[<?php echo esc_attr( $rule_type ); ?>][][custom][week][day][]" data-field="custom-week-day" value="7" {{tribe_checked_if_in '7' custom.week.day}}/>
		<?php esc_html_e( 'Su', 'tribe-events-calendar-pro' ); ?>
	</label>
</div>
