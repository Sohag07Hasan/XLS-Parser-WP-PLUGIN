<div class="wrap">
	<h2> Products Upload With XLS </h2>
	
	<?php 
		echo self::print_message();
	?>
	
	<p>Please use the prescribed <a href="<?php echo self::get_demo_xls(); ?>"> format </a> </p>
	
	
	<form action="" method="post" enctype="multipart/form-data" >
		<input type="hidden" name="producto-xls-importer-form" value="Y" />
		
		<table class="form-table">
			<tr>
				<td> <input type="file" name="producto-xls-file" /> <input type="submit" value="import" class="button-primary" /></td>
			</tr>
		</table>
	
	</form>
</div>