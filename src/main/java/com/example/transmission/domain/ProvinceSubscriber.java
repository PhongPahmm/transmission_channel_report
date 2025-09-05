package com.example.transmission.domain;

import javax.persistence.*;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.AllArgsConstructor;

@Entity
@Table(name = "rp_province_subscribers")
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ProvinceSubscriber {
    
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;
    
    @Column(name = "province_name", length = 100)
    private String provinceName;
    
    @Column(name = "total_subscribers")
    private Long totalSubscribers;
}
