package com.example.transmission;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.TestPropertySource;
import com.example.transmission.service.TransmissionChannelIncidentService;
import java.time.LocalDate;

@SpringBootTest
@TestPropertySource(properties = {
    "spring.datasource.url=jdbc:h2:mem:testdb",
    "spring.datasource.driver-class-name=org.h2.Driver",
    "spring.jpa.hibernate.ddl-auto=create-drop"
})
class TransmissionApplicationTests {

	@Autowired
	private TransmissionChannelIncidentService transmissionChannelIncidentService;

	@Test
	void contextLoads() {
	}

	@Test
	void testExportExcelFileWithNullHandling() {
		// Test that the service handles null data gracefully
		LocalDate testDate = LocalDate.now();
		try {
			// This should not throw a NullPointerException even if data is null
			transmissionChannelIncidentService.exportExcelFile(testDate);
		} catch (Exception e) {
			// We expect some exceptions due to missing data, but not NullPointerException
			assert !(e instanceof NullPointerException) : "NullPointerException should be handled: " + e.getMessage();
		}
	}

}
